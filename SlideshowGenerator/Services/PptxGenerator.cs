using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using SlideshowGenerator.Models;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace LanguageCourseSlides.Services;

public static class PptxGenerator
{
    // -----------------------------------------------------------------------
    // Public entry point
    // -----------------------------------------------------------------------

    public static void Generate(
        TemplateConfig config,
        string outputPath,
        List<WordEntry> entries,
        IProgress<int>? progress = null)
    {
        if (!File.Exists(config.TemplatePath))
            throw new FileNotFoundException("Template file not found.", config.TemplatePath);

        File.Copy(config.TemplatePath, outputPath, overwrite: true);

        using var prs = PresentationDocument.Open(outputPath, isEditable: true);
        var presPart = prs.PresentationPart!;
        var presentation = presPart.Presentation;
        var slideIdList = presentation.SlideIdList!;

        // Snapshot original template slides before we start cloning
        var templateIds = slideIdList.Elements<SlideId>().ToList();
        var templateParts = templateIds
            .Select(id => (SlidePart)presPart.GetPartById(id.RelationshipId!))
            .ToList();

        // ── PASS 1: Clone every output slide in order ──────────────────────
        var output = new List<(SlidePart Part, SlideRole Role, WordEntry? Entry)>();
        int done = 0;

        foreach (var def in config.Slides.OrderBy(s => s.SlideIndex))
        {
            if (def.SlideIndex >= templateParts.Count) continue;
            var source = templateParts[def.SlideIndex];

            switch (def.Role)
            {
                case SlideRole.Static:
                    output.Add((CloneSlide(presPart, source), SlideRole.Static, null));
                    break;

                case SlideRole.Index:
                    // Index content injected in Pass 2 once slide positions are known
                    output.Add((CloneSlide(presPart, source), SlideRole.Index, null));
                    break;

                case SlideRole.Word:
                    foreach (var entry in entries)
                    {
                        var wordPart = CloneSlide(presPart, source);
                        ProcessWordSlide(wordPart, entry, config);
                        output.Add((wordPart, SlideRole.Word, entry));
                        progress?.Report(++done);
                    }
                    break;
            }
        }

        // ── Remove original template slides; register new ones ─────────────
        foreach (var id in templateIds) slideIdList.RemoveChild(id);
        foreach (var part in templateParts) presPart.DeletePart(part);

        uint sid = 256;
        foreach (var (part, _, _) in output)
        {
            var relId = presPart.GetIdOfPart(part);
            slideIdList.Append(new SlideId { Id = sid++, RelationshipId = relId });
        }

        // ── PASS 2: Build index now that slide positions are known ─────────
        var indexDef = config.IndexSlide;
        if (indexDef != null)
        {
            // Map every WordEntry → its 1-based slide number in the output
            var wordSlideMap = new Dictionary<WordEntry, int>();
            int pos = 1;
            foreach (var (_, role, entry) in output)
            {
                if (role == SlideRole.Word && entry != null)
                    wordSlideMap[entry] = pos;
                pos++;
            }

            foreach (var (indexPart, role, _) in output.Where(o => o.Role == SlideRole.Index))
            {
                BuildIndex(
                    indexPart, entries, wordSlideMap,
                    indexDef.IndexPlaceholder,
                    config.IndexLineFormat,
                    config.HyperlinkIndex);
            }
        }

        presentation.Save();
    }

    // =======================================================================
    // Word slide processing
    // =======================================================================

    private static void ProcessWordSlide(
        SlidePart slidePart,
        WordEntry entry,
        TemplateConfig config)
    {
        MergeTextRuns(slidePart);
        SubstitutePlaceholders(slidePart, entry.ToPlaceholders());

        if (entry.HasImage)
            InjectImage(slidePart, entry.Image!, config.ImagePlaceholder);
        else
            RemovePlaceholderShape(slidePart, config.ImagePlaceholder);

        if (entry.HasAudio)
            InjectAudio(slidePart, entry.Audio!, config.AudioPlaceholder);
        else
            RemovePlaceholderShape(slidePart, config.AudioPlaceholder);

        slidePart.Slide.Save();
    }

    // =======================================================================
    // Index builder — pure DOM manipulation, no raw XML replacement
    // =======================================================================

    private static void BuildIndex(
        SlidePart indexPart,
        List<WordEntry> entries,
        Dictionary<WordEntry, int> wordSlideMap,
        string placeholder,
        string lineFormat,
        bool hyperlink)
    {
        // Ensure any prior raw-stream writes are flushed before DOM access
        indexPart.Slide.Save();

        // MergeTextRuns so the placeholder token isn't split across runs
        MergeTextRuns(indexPart);

        var slide = indexPart.Slide;

        // Find the paragraph containing {{Index}} (or whatever placeholder)
        var targetPara = slide.Descendants<A.Paragraph>()
            .FirstOrDefault(p => p.InnerText.Contains(placeholder));

        if (targetPara == null)
            return;   // placeholder not present on this slide — skip silently

        var txBody = targetPara.Ancestors<A.TextBody>().FirstOrDefault();
        if (txBody == null) return;

        // Capture formatting from the placeholder paragraph to reuse on entries
        var templateRun = targetPara.Elements<A.Run>().FirstOrDefault();
        var templatePPr = targetPara
            .Elements<A.ParagraphProperties>()
            .FirstOrDefault()
            ?.CloneNode(true) as A.ParagraphProperties;

        // Remove the {{Index}} placeholder paragraph
        targetPara.Remove();

        // Insert one paragraph per entry
        for (int i = 0; i < entries.Count; i++)
        {
            var entry = entries[i];
            var lineText = FormatLine(lineFormat, i + 1, entry);

            var para = new A.Paragraph();

            // Copy paragraph-level formatting (indent, spacing etc.)
            if (templatePPr != null)
                para.Append((A.ParagraphProperties)templatePPr.CloneNode(true));

            var rPr = CloneRunProperties(templateRun);
            rPr.Dirty = false;

            if (hyperlink && wordSlideMap.TryGetValue(entry, out int targetSlide))
            {
                // Add internal hyperlink relationship
                var rel = indexPart.AddHyperlinkRelationship(
                    new Uri($"slide{targetSlide}.xml", UriKind.Relative),
                    isExternal: false);

                // hlinkClick must be first child of rPr
                rPr.InsertAt(new A.HyperlinkOnClick { Id = rel.Id }, 0);
            }

            var run = new A.Run();
            run.Append(rPr);
            run.Append(new A.Text(lineText));

            para.Append(run);
            txBody.Append(para);
        }

        // Flush DOM back to the part stream
        indexPart.Slide.Save();
    }

    private static A.RunProperties CloneRunProperties(A.Run? source)
    {
        if (source?.RunProperties != null)
            return (A.RunProperties)source.RunProperties.CloneNode(true);

        return new A.RunProperties { Language = "en-US", Dirty = false };
    }

    // =======================================================================
    // Shape removal
    // =======================================================================

    private static void RemovePlaceholderShape(SlidePart slidePart, string placeholder)
    {
        var slide = slidePart.Slide;

        slide.Descendants<P.Shape>()
             .Where(sp => sp.InnerText.Contains(placeholder))
             .ToList().ForEach(sp => sp.Remove());

        slide.Descendants<P.Picture>()
             .Where(pic => pic.InnerText.Contains(placeholder))
             .ToList().ForEach(pic => pic.Remove());

        slide.Descendants<P.GraphicFrame>()
             .Where(gf => gf.InnerText.Contains(placeholder))
             .ToList().ForEach(gf => gf.Remove());
    }

    // =======================================================================
    // Image injection
    // =======================================================================

    private static void InjectImage(SlidePart slidePart, AssetData asset, string placeholder)
    {
        var slide = slidePart.Slide;

        var placeholderShape = slide.Descendants<P.Shape>()
            .FirstOrDefault(sp => sp.InnerText.Contains(placeholder));

        if (placeholderShape == null) return;

        var xfrm = placeholderShape.ShapeProperties?.Transform2D;
        if (xfrm == null) return;

        long x = xfrm.Offset?.X ?? 0;
        long y = xfrm.Offset?.Y ?? 0;
        long cx = xfrm.Extents?.Cx ?? 1_000_000;
        long cy = xfrm.Extents?.Cy ?? 1_000_000;

        var imagePart = slidePart.AddImagePart(asset.ContentType);
        using (var stream = asset.OpenStream())
            imagePart.FeedData(stream);

        var relId = slidePart.GetIdOfPart(imagePart);
        var pic = BuildPicture(relId, "img", x, y, cx, cy);

        (placeholderShape.Parent as P.ShapeTree)?.InsertAfter(pic, placeholderShape);
        placeholderShape.Remove();
    }

    private static P.Picture BuildPicture(
        string relId, string name,
        long x, long y, long cx, long cy)
    {
        var pic = new P.Picture();

        pic.Append(new P.NonVisualPictureProperties(
            new P.NonVisualDrawingProperties { Id = 100, Name = name },
            new P.NonVisualPictureDrawingProperties(
                new A.PictureLocks { NoChangeAspect = true }),
            new ApplicationNonVisualDrawingProperties()));

        pic.Append(new P.BlipFill(
            new A.Blip { Embed = relId },
            new A.Stretch(new A.FillRectangle())));

        pic.Append(new P.ShapeProperties(
            new A.Transform2D(
                new A.Offset { X = x, Y = y },
                new A.Extents { Cx = cx, Cy = cy }),
            new A.PresetGeometry(new A.AdjustValueList())
            { Preset = A.ShapeTypeValues.Rectangle }));

        return pic;
    }

    // =======================================================================
    // Audio injection
    // =======================================================================

    private static void InjectAudio(SlidePart slidePart, AssetData asset, string placeholder)
    {
        var slide = slidePart.Slide;

        var placeholderShape = slide.Descendants<P.Shape>()
            .FirstOrDefault(sp => sp.InnerText.Contains(placeholder));

        long x = 914_400, y = 4_572_000;
        long cx = 457_200, cy = 457_200;

        if (placeholderShape?.ShapeProperties?.Transform2D is { } xfrm)
        {
            x = xfrm.Offset?.X ?? x;
            y = xfrm.Offset?.Y ?? y;
            cx = xfrm.Extents?.Cx ?? cx;
            cy = xfrm.Extents?.Cy ?? cy;
        }

        MediaDataPart mediaPart;
        try
        {
            mediaPart = slidePart.OpenXmlPackage.CreateMediaDataPart(
                asset.ContentType, asset.Extension);
        }
        catch
        {
            mediaPart = slidePart.OpenXmlPackage.CreateMediaDataPart(
                "audio/mpeg", ".mp3");
        }

        using (var s = asset.OpenStream())
            mediaPart.FeedData(s);

        var audioRelId = slidePart.AddAudioReferenceRelationship(mediaPart).Id;
        var mediaRelId = slidePart.AddMediaReferenceRelationship(mediaPart).Id;

        var audioShape = BuildAudioShape(audioRelId, mediaRelId,
            System.IO.Path.GetFileNameWithoutExtension(asset.FilePath ?? "audio"),
            x, y, cx, cy);

        slide.CommonSlideData?.ShapeTree?.Append(audioShape);
        placeholderShape?.Remove();
    }

    private static P.Picture BuildAudioShape(
        string audioRelId, string mediaRelId, string name,
        long x, long y, long cx, long cy)
    {
        var pic = new P.Picture();
        var nvPPr = new P.NonVisualPictureProperties();

        var cNvPr = new P.NonVisualDrawingProperties { Id = 200, Name = name };
        cNvPr.Append(new A.HyperlinkOnClick
        {
            Id = audioRelId,
            Action = "ppaction://media"
        });
        nvPPr.Append(cNvPr);
        nvPPr.Append(new P.NonVisualPictureDrawingProperties());

        var appNvPr = new ApplicationNonVisualDrawingProperties();
        appNvPr.Append(new AudioFromFile { Link = mediaRelId });
        nvPPr.Append(appNvPr);

        pic.Append(nvPPr);
        pic.Append(new P.BlipFill(
            new A.Blip(),
            new A.Stretch(new A.FillRectangle())));
        pic.Append(new P.ShapeProperties(
            new A.Transform2D(
                new A.Offset { X = x, Y = y },
                new A.Extents { Cx = cx, Cy = cy }),
            new A.PresetGeometry(new A.AdjustValueList())
            { Preset = A.ShapeTypeValues.Rectangle }));

        return pic;
    }

    // =======================================================================
    // Helpers
    // =======================================================================

    private static SlidePart CloneSlide(PresentationPart presPart, SlidePart source)
    {
        var newPart = presPart.AddNewPart<SlidePart>();
        using (var stream = source.GetStream())
            newPart.FeedData(stream);
        foreach (var rel in source.Parts)
            newPart.AddPart(rel.OpenXmlPart, rel.RelationshipId);
        return newPart;
    }

    /// <summary>
    /// Raw XML string-replace for word-slide text placeholders only.
    /// Values are XML-escaped. Never called for the index slide.
    /// </summary>
    private static void SubstitutePlaceholders(
        SlidePart slidePart, Dictionary<string, string> map)
    {
        string xml;
        using (var r = new StreamReader(slidePart.GetStream()))
            xml = r.ReadToEnd();

        foreach (var (key, value) in map)
            xml = xml.Replace(key, System.Security.SecurityElement.Escape(value ?? ""));

        using var w = new StreamWriter(slidePart.GetStream(FileMode.Create));
        w.Write(xml);
    }

    /// <summary>
    /// Merge runs within any paragraph that contains {{ so a placeholder
    /// split across multiple runs is consolidated before DOM search/replace.
    /// </summary>
    private static void MergeTextRuns(SlidePart slidePart)
    {
        var slide = slidePart.Slide;

        foreach (var para in slide.Descendants<A.Paragraph>().ToList())
        {
            var runs = para.Elements<A.Run>().ToList();
            if (runs.Count < 2) continue;

            var combined = string.Concat(runs.Select(r => r.Text?.Text ?? ""));
            if (!combined.Contains("{{")) continue;

            runs[0].Text = new A.Text(combined);

            for (int i = 1; i < runs.Count; i++)
                runs[i].Remove();
        }
    }

    private static string FormatLine(string fmt, int n, WordEntry e) =>
        fmt.Replace("{n}", n.ToString())
           .Replace("{word}", e.Word)
           .Replace("{plural}", e.Plural)
           .Replace("{type}", e.Type)
           .Replace("{english}", e.English);
}
