using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using LanguageCourseSlides.Models;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace LanguageCourseSlides.Services;

public static class PptxGenerator
{
    // -----------------------------------------------------------------------
    // Shape name constants — these must match names in the .pptx template
    // -----------------------------------------------------------------------
    public const string ShapeAudio = "shape_audio";
    public const string ShapeImage = "shape_image";
    public const string ShapeNext  = "shape_next";
    public const string ShapeIndex = "shape_index";

    // -----------------------------------------------------------------------
    // Public entry point
    // -----------------------------------------------------------------------

    public static void Generate(
        TemplateConfig   config,
        string           outputPath,
        List<WordEntry>  entries,
        IProgress<int>?  progress = null)
    {
        if (!File.Exists(config.TemplatePath))
            throw new FileNotFoundException("Template file not found.", config.TemplatePath);

        File.Copy(config.TemplatePath, outputPath, overwrite: true);

        using var doc    = PresentationDocument.Open(outputPath, isEditable: true);
        var presPart     = doc.PresentationPart!;
        var presentation = presPart.Presentation;
        var slideIdList  = presentation.SlideIdList!;

        // Snapshot original template slides
        var templateIds   = slideIdList.Elements<SlideId>().ToList();
        var templateParts = templateIds
            .Select(id => (SlidePart)presPart.GetPartById(id.RelationshipId!))
            .ToList();

        // ── PASS 1: Build output slide list ───────────────────────────────
        // Each item: (SlidePart, SlideRole, WordEntry?, indexPageEntries?)
        var output = new List<OutputSlide>();
        int done   = 0;

        // How many index pages do we need?
        int wordsPerPage = Math.Max(1, config.WordsPerIndexPage);

        foreach (var def in config.Slides.OrderBy(s => s.SlideIndex))
        {
            if (def.SlideIndex >= templateParts.Count) continue;
            var source = templateParts[def.SlideIndex];

            switch (def.Role)
            {
                case SlideRole.Static:
                    output.Add(new OutputSlide(CloneSlide(presPart, source), SlideRole.Static));
                    break;

                case SlideRole.Index:
                    // One index slide per batch of wordsPerPage entries
                    int totalIndexPages = Math.Max(1, (int)Math.Ceiling(entries.Count / (double)wordsPerPage));
                    for (int page = 0; page < totalIndexPages; page++)
                    {
                        var batch = entries
                            .Skip(page * wordsPerPage)
                            .Take(wordsPerPage)
                            .ToList();
                        output.Add(new OutputSlide(CloneSlide(presPart, source), SlideRole.Index)
                        {
                            IndexEntries = batch,
                            IndexPage    = page + 1,
                        });
                    }
                    break;

                case SlideRole.Word:
                    foreach (var entry in entries)
                    {
                        var wordPart = CloneSlide(presPart, source);
                        ProcessWordSlide(wordPart, entry, config);
                        output.Add(new OutputSlide(wordPart, SlideRole.Word) { Entry = entry });
                        progress?.Report(++done);
                    }
                    break;
            }
        }

        // ── Remove old template slides; register new ones ─────────────────
        foreach (var id   in templateIds)   slideIdList.RemoveChild(id);
        foreach (var part in templateParts) presPart.DeletePart(part);

        uint sid = 256;
        for (int i = 0; i < output.Count; i++)
        {
            var relId = presPart.GetIdOfPart(output[i].Part);
            slideIdList.Append(new SlideId { Id = sid++, RelationshipId = relId });
        }

        // ── Name each slide (sets <p:cSld name="...">) ────────────────────
        for (int i = 0; i < output.Count; i++)
        {
            var item = output[i];
            string name = item.Role switch
            {
                SlideRole.Static => $"Slide_{i + 1:000}_static",
                SlideRole.Index  => output.Count(o => o.Role == SlideRole.Index) > 1
                                     ? $"Slide_Index_{item.IndexPage}"
                                     : "Slide_Index",
                SlideRole.Word   => $"Slide_{SanitizeName(item.Entry!.Word)}",
                _                => $"Slide_{i + 1:000}",
            };
            SetSlideName(item.Part, name);
        }

        // ── PASS 2: Populate index slides & wire navigation shapes ─────────

        // First index slide position (1-based) — shape_index links here
        int firstIndexPos = output.FindIndex(o => o.Role == SlideRole.Index) + 1;

        // Build word-entry → 1-based slide position map
        var wordSlidePos = new Dictionary<WordEntry, int>();
        for (int i = 0; i < output.Count; i++)
            if (output[i].Role == SlideRole.Word && output[i].Entry != null)
                wordSlidePos[output[i].Entry!] = i + 1;

        for (int i = 0; i < output.Count; i++)
        {
            var item = output[i];

            if (item.Role == SlideRole.Index)
            {
                BuildIndexContent(
                    item.Part,
                    item.IndexEntries,
                    wordSlidePos,
                    config.IndexLineFormat,
                    config.HyperlinkIndex);
            }
            else if (item.Role == SlideRole.Word)
            {
                int nextPos  = i + 2; // next slide (1-based), wraps if last
                if (nextPos > output.Count) nextPos = firstIndexPos > 0 ? firstIndexPos : 1;

                WireNavigationShapes(item.Part, nextPos, firstIndexPos, output.Count);
            }
        }

        presentation.Save();
    }

    // =======================================================================
    // Word slide — text substitution + asset handling
    // =======================================================================

    private static void ProcessWordSlide(
        SlidePart      slidePart,
        WordEntry      entry,
        TemplateConfig config)
    {
        // 1. Merge split runs so {{placeholders}} are intact strings
        MergeTextRuns(slidePart);

        // 2. Replace all {{Field}} text placeholders via raw XML
        SubstitutePlaceholders(slidePart, entry.ToPlaceholders());

        // 3. Image — find by shape name
        if (entry.HasImage)
            InjectImage(slidePart, entry.Image!, ShapeImage);
        else
            RemoveShapeByName(slidePart, ShapeImage);

        // 4. Audio — find by shape name
        if (entry.HasAudio)
            InjectAudio(slidePart, entry.Audio!, ShapeAudio);
        else
            RemoveShapeByName(slidePart, ShapeAudio);

        // shape_next and shape_index wired in Pass 2 once slide order is known
        slidePart.Slide.Save();
    }

    // =======================================================================
    // Navigation shapes on word slides (Pass 2)
    // =======================================================================

    private static void WireNavigationShapes(
        SlidePart slidePart,
        int       nextSlidePos,
        int       indexSlidePos,
        int       totalSlides)
    {
        // shape_next → hyperlink to next slide
        var nextShape = FindShapeByName(slidePart, ShapeNext);
        if (nextShape != null && nextSlidePos >= 1 && nextSlidePos <= totalSlides)
        {
            var rel = slidePart.AddHyperlinkRelationship(
                new Uri($"slide{nextSlidePos}.xml", UriKind.Relative), false);
            EnsureClickAction(nextShape, rel.Id);
        }

        // shape_index → hyperlink back to index
        var idxShape = FindShapeByName(slidePart, ShapeIndex);
        if (idxShape != null && indexSlidePos >= 1)
        {
            var rel = slidePart.AddHyperlinkRelationship(
                new Uri($"slide{indexSlidePos}.xml", UriKind.Relative), false);
            EnsureClickAction(idxShape, rel.Id);
        }

        slidePart.Slide.Save();
    }

    private static void EnsureClickAction(P.Shape shape, string relId)
    {
        // Add a hlinkClick to every run in the shape, or to the shape's spPr
        // The easiest approach: add to NonVisualShapeDrawingProperties
        var nvSpPr = shape.NonVisualShapeProperties;
        if (nvSpPr == null) return;

        var cNvPr = nvSpPr.NonVisualDrawingProperties;
        if (cNvPr == null) return;

        // Remove any existing hyperlink on click first
        var existing = cNvPr.Elements<A.HyperlinkOnClick>().FirstOrDefault();
        existing?.Remove();

        cNvPr.Append(new A.HyperlinkOnClick { Id = relId });
    }

    // =======================================================================
    // Index content builder — pure DOM, no raw XML replacement
    // =======================================================================

    private static void BuildIndexContent(
        SlidePart                  indexPart,
        List<WordEntry>            entries,
        Dictionary<WordEntry, int> wordSlidePos,
        string                     lineFormat,
        bool                       hyperlink)
    {
        var slide = indexPart.Slide;

        // Merge split runs first
        MergeTextRuns(indexPart);

        // Find the paragraph containing {{Index}} in any text body on this slide
        A.Paragraph? targetPara = null;
        A.TextBody?  txBody     = null;

        foreach (var para in slide.Descendants<A.Paragraph>())
        {
            if (para.InnerText.Contains("{{Index}}"))
            {
                targetPara = para;
                txBody     = para.Ancestors<A.TextBody>().FirstOrDefault();
                break;
            }
        }

        if (targetPara == null || txBody == null)
        {
            // Fallback: search every text body for any {{…}} and dump raw text
            // (template may not have {{Index}} — nothing we can do without it)
            return;
        }

        // Capture run/paragraph formatting from the placeholder paragraph
        var templateRun = targetPara.Elements<A.Run>().FirstOrDefault();
        var templatePPr = targetPara
            .Elements<A.ParagraphProperties>()
            .FirstOrDefault()
            ?.CloneNode(true) as A.ParagraphProperties;

        // Remove placeholder paragraph
        targetPara.Remove();

        // Insert one paragraph per entry
        for (int i = 0; i < entries.Count; i++)
        {
            var entry    = entries[i];
            var lineText = FormatLine(lineFormat, i + 1, entry);

            var para = new A.Paragraph();
            if (templatePPr != null)
                para.Append((A.ParagraphProperties)templatePPr.CloneNode(true));

            var rPr = CloneRunProperties(templateRun);
            rPr.Dirty = false;

            if (hyperlink && wordSlidePos.TryGetValue(entry, out int targetPos))
            {
                var rel = indexPart.AddHyperlinkRelationship(
                    new Uri($"slide{targetPos}.xml", UriKind.Relative), false);
                rPr.InsertAt(new A.HyperlinkOnClick { Id = rel.Id }, 0);
            }

            var run = new A.Run();
            run.Append(rPr);
            run.Append(new A.Text(lineText));
            para.Append(run);
            txBody.Append(para);
        }

        indexPart.Slide.Save();
    }

    // =======================================================================
    // Shape finder by PowerPoint shape name (<p:cNvPr name="...">)
    // =======================================================================

    private static P.Shape? FindShapeByName(SlidePart slidePart, string name)
    {
        return slidePart.Slide
            .Descendants<P.Shape>()
            .FirstOrDefault(sp =>
                string.Equals(
                    sp.NonVisualShapeProperties
                      ?.NonVisualDrawingProperties
                      ?.Name?.Value,
                    name,
                    StringComparison.OrdinalIgnoreCase));
    }

    private static void RemoveShapeByName(SlidePart slidePart, string name)
    {
        var shape = FindShapeByName(slidePart, name);
        shape?.Remove();
    }

    // =======================================================================
    // Image injection — replaces shape_image with the actual picture
    // =======================================================================

    private static void InjectImage(SlidePart slidePart, AssetData asset, string shapeName)
    {
        var shape = FindShapeByName(slidePart, shapeName);
        if (shape == null) return;

        var xfrm = shape.ShapeProperties?.Transform2D;
        if (xfrm == null) return;

        long x  = xfrm.Offset?.X  ?? 0;
        long y  = xfrm.Offset?.Y  ?? 0;
        long cx = xfrm.Extents?.Cx ?? 1_000_000;
        long cy = xfrm.Extents?.Cy ?? 1_000_000;

        var imagePart = slidePart.AddImagePart(asset.ContentType);
        using (var stream = asset.OpenStream())
            imagePart.FeedData(stream);

        var relId = slidePart.GetIdOfPart(imagePart);
        var pic   = BuildPicture(relId, shapeName, x, y, cx, cy);

        (shape.Parent as P.ShapeTree)?.InsertAfter(pic, shape);
        shape.Remove();
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
                new A.Offset  { X = x,  Y = y  },
                new A.Extents { Cx = cx, Cy = cy }),
            new A.PresetGeometry(new A.AdjustValueList())
                { Preset = A.ShapeTypeValues.Rectangle }));
        return pic;
    }

    // =======================================================================
    // Audio injection — attaches media to shape_audio
    // =======================================================================

    private static void InjectAudio(SlidePart slidePart, AssetData asset, string shapeName)
    {
        var shape = FindShapeByName(slidePart, shapeName);

        // Default position if shape not found
        long x  = 457_200, y  = 457_200;
        long cx = 457_200, cy = 457_200;

        if (shape?.ShapeProperties?.Transform2D is { } xfrm)
        {
            x  = xfrm.Offset?.X  ?? x;
            y  = xfrm.Offset?.Y  ?? y;
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
            mediaPart = slidePart.OpenXmlPackage.CreateMediaDataPart("audio/mpeg", ".mp3");
        }

        using (var s = asset.OpenStream())
            mediaPart.FeedData(s);

        var audioRelId = slidePart.AddAudioReferenceRelationship(mediaPart).Id;
        var mediaRelId = slidePart.AddMediaReferenceRelationship(mediaPart).Id;

        var audioShape = BuildAudioShape(
            audioRelId, mediaRelId,
            shapeName, x, y, cx, cy);

        slidePart.Slide.CommonSlideData?.ShapeTree?.Append(audioShape);
        shape?.Remove();
    }

    private static P.Picture BuildAudioShape(
        string audioRelId, string mediaRelId, string name,
        long x, long y, long cx, long cy)
    {
        var pic   = new P.Picture();
        var nvPPr = new P.NonVisualPictureProperties();

        var cNvPr = new P.NonVisualDrawingProperties { Id = 200, Name = name };
        cNvPr.Append(new A.HyperlinkOnClick { Id = audioRelId, Action = "ppaction://media" });
        nvPPr.Append(cNvPr);
        nvPPr.Append(new P.NonVisualPictureDrawingProperties());

        var appNvPr = new ApplicationNonVisualDrawingProperties();
        appNvPr.Append(new AudioFromFile { Link = mediaRelId });
        nvPPr.Append(appNvPr);

        pic.Append(nvPPr);
        pic.Append(new P.BlipFill(new A.Blip(), new A.Stretch(new A.FillRectangle())));
        pic.Append(new P.ShapeProperties(
            new A.Transform2D(
                new A.Offset  { X = x,  Y = y  },
                new A.Extents { Cx = cx, Cy = cy }),
            new A.PresetGeometry(new A.AdjustValueList())
                { Preset = A.ShapeTypeValues.Rectangle }));
        return pic;
    }

    // =======================================================================
    // Slide naming
    // =======================================================================

    private static void SetSlideName(SlidePart slidePart, string name)
    {
        var cSld = slidePart.Slide.CommonSlideData;
        if (cSld != null) cSld.Name = name;
    }

    private static string SanitizeName(string word)
    {
        var sb = new StringBuilder();
        foreach (char c in word)
            sb.Append(char.IsLetterOrDigit(c) ? c : '_');
        return sb.ToString().Trim('_');
    }

    // =======================================================================
    // Shared helpers
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
    /// Raw XML string-replace for {{Field}} text placeholders on word slides only.
    /// Values are XML-escaped. Not used for index content.
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
    /// Merge runs within paragraphs that contain {{ so a split placeholder
    /// becomes a single run before DOM search/replace.
    /// </summary>
    private static void MergeTextRuns(SlidePart slidePart)
    {
        foreach (var para in slidePart.Slide.Descendants<A.Paragraph>().ToList())
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

    private static A.RunProperties CloneRunProperties(A.Run? source)
    {
        if (source?.RunProperties != null)
            return (A.RunProperties)source.RunProperties.CloneNode(true);
        return new A.RunProperties { Language = "en-US", Dirty = false };
    }

    private static string FormatLine(string fmt, int n, WordEntry e) =>
        fmt.Replace("{n}",       n.ToString())
           .Replace("{word}",    e.Word)
           .Replace("{plural}",  e.Plural)
           .Replace("{type}",    e.Type)
           .Replace("{english}", e.English);

    // =======================================================================
    // Internal data carrier
    // =======================================================================

    private class OutputSlide
    {
        public SlidePart       Part         { get; }
        public SlideRole       Role         { get; }
        public WordEntry?      Entry        { get; init; }
        public List<WordEntry> IndexEntries { get; init; } = [];
        public int             IndexPage    { get; init; } = 1;

        public OutputSlide(SlidePart part, SlideRole role)
        {
            Part = part;
            Role = role;
        }
    }
}
