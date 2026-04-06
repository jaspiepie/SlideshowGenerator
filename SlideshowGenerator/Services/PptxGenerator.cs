using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using LanguageCourseSlides.Models;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace LanguageCourseSlides.Services;

public static class PptxGenerator
{
    // -----------------------------------------------------------------------
    // Shape name constants — must match names in the .pptx template
    // -----------------------------------------------------------------------
    public const string ShapeAudio = "shape_audio";
    public const string ShapeImage = "shape_image";
    public const string ShapeNext  = "shape_next";
    public const string ShapeIndex = "shape_index";

    // -----------------------------------------------------------------------
    // PowerPoint requires this action string on every internal slide hyperlink.
    // Without it, the click is registered but PowerPoint does not jump slides.
    // The relationship type is the standard "hyperlink" with isExternal=false
    // (TargetMode="Internal") — this is exactly what PowerPoint itself writes.
    // -----------------------------------------------------------------------
    private const string SlideJumpAction = "ppaction://hlinksldjump";

    // -----------------------------------------------------------------------
    // Column definitions — token → (header label, relative width weight)
    // -----------------------------------------------------------------------
    private static readonly IReadOnlyDictionary<string, (string Header, int Weight)> ColDefs =
        new Dictionary<string, (string, int)>(StringComparer.OrdinalIgnoreCase)
        {
            ["{n}"]       = ("#",       1),
            ["{word}"]    = ("Word",    4),
            ["{plural}"]  = ("Plural",  3),
            ["{type}"]    = ("Type",    2),
            ["{english}"] = ("English", 4),
        };

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
        var slideIdList  = presentation!.SlideIdList!;

        var templateIds   = slideIdList.Elements<SlideId>().ToList();
        var templateParts = templateIds
            .Select(id => (SlidePart)presPart.GetPartById(id.RelationshipId!))
            .ToList();

        var columns = ParseColumns(config.IndexLineFormat);

        // ── PASS 1: Clone all output slides ──────────────────────────────
        var output       = new List<OutputSlide>();
        int done         = 0;
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
                    int totalPages = Math.Max(1,
                        (int)Math.Ceiling(entries.Count / (double)wordsPerPage));
                    for (int page = 0; page < totalPages; page++)
                    {
                        var batch = entries
                            .Skip(page * wordsPerPage)
                            .Take(wordsPerPage)
                            .ToList();
                        output.Add(new OutputSlide(CloneSlide(presPart, source), SlideRole.Index)
                        {
                            IndexEntries    = batch,
                            IndexPage       = page + 1,
                            GlobalRowOffset = page * wordsPerPage,
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

        // ── Remove template slides; register new ones ─────────────────────
        foreach (var id   in templateIds)   slideIdList.RemoveChild(id);
        foreach (var part in templateParts) presPart.DeletePart(part);

        uint sid = 256;
        foreach (var item in output)
        {
            var relId = presPart.GetIdOfPart(item.Part);
            slideIdList.Append(new SlideId { Id = sid++, RelationshipId = relId });
        }

        // ── Name every slide ──────────────────────────────────────────────
        bool multiIndex = output.Count(o => o.Role == SlideRole.Index) > 1;
        for (int i = 0; i < output.Count; i++)
        {
            var item = output[i];
            string name = item.Role switch
            {
                SlideRole.Static => $"Slide_{i + 1:000}_static",
                SlideRole.Index  => multiIndex
                                    ? $"Slide_Index_{item.IndexPage}"
                                    : "Slide_Index",
                SlideRole.Word   => $"Slide_{SanitizeName(item.Entry!.Word)}",
                _                => $"Slide_{i + 1:000}",
            };
            SetSlideName(item.Part, name);
            item.Part.Slide.Save();
        }

        // ── PASS 2A: Build index tables ───────────────────────────────────
        // Runs after every output slide exists so index hyperlinks can
        // resolve against the final slide parts.

        var wordPartMap = new Dictionary<WordEntry, SlidePart>();
        foreach (var item in output)
            if (item.Role == SlideRole.Word && item.Entry != null)
                wordPartMap[item.Entry] = item.Part;

        SlidePart? firstIndexPart = output
            .FirstOrDefault(o => o.Role == SlideRole.Index)?.Part;

        foreach (var item in output)
        {
            if (item.Role == SlideRole.Index)
            {
                BuildIndexTable(
                    item.Part, item.IndexEntries, wordPartMap,
                    columns, item.GlobalRowOffset, config.HyperlinkIndex);
            }
        }

        // ── PASS 2B: Wire navigation buttons ─────────────────────────────
        // Runs last so shape_next and shape_index are resolved against
        // the final registered output slides.

        for (int i = 0; i < output.Count; i++)
        {
            var item = output[i];
            if (item.Role == SlideRole.Word)
            {
                SlidePart? nextPart = (i + 1 < output.Count)
                    ? output[i + 1].Part
                    : firstIndexPart;

                WireNavigationShapes(item.Part, nextPart, firstIndexPart);
            }
        }

        presentation.Save();
    }

    private static string AddSlideJumpRelationship(SlidePart source, SlidePart target)
    {
        var targetUri = BuildSlideJumpUri(target);

        foreach (var rel in source.HyperlinkRelationships)
        {
            if (rel.IsExternal) continue;
            if (rel.Uri == targetUri)
                return rel.Id;
        }

        return source.AddHyperlinkRelationship(targetUri, isExternal: false).Id;
    }

    // =======================================================================
    // Column parser
    // =======================================================================

    private static List<string> ParseColumns(string format)
    {
        var tokens = Regex.Matches(format, @"\{[a-z]+\}", RegexOptions.IgnoreCase)
            .Select(m => m.Value.ToLowerInvariant())
            .Where(t => ColDefs.ContainsKey(t))
            .Distinct()
            .ToList();

        if (tokens.Count == 0) tokens.Add("{word}");
        return tokens;
    }

    // =======================================================================
    // Index table builder
    // =======================================================================

    private static void BuildIndexTable(
        SlidePart                        indexPart,
        List<WordEntry>                  entries,
        Dictionary<WordEntry, SlidePart> wordPartMap,
        List<string>                     columns,
        int                              rowOffset,
        bool                             hyperlink)
    {
        var slide = indexPart.Slide;

        P.Shape? placeholder = FindShapeContainingText(slide, "{{Index}}");
        if (placeholder == null) return;

        var xfrm = placeholder.ShapeProperties?.Transform2D;
        long x  = xfrm?.Offset?.X  ?? 457_200L;
        long y  = xfrm?.Offset?.Y  ?? 914_400L;
        long cx = xfrm?.Extents?.Cx ?? 8_229_600L;
        long cy = xfrm?.Extents?.Cy ?? 3_657_600L;

        placeholder.Remove();

        int  totalWeight = columns.Sum(c => ColDefs[c].Weight);
        var  colWidths   = columns
            .Select(c => (long)(cx * ColDefs[c].Weight / (double)totalWeight))
            .ToList();

        long hdrH = 457_200L;
        long rowH = entries.Count > 0
            ? Math.Max(304_800L, (cy - hdrH) / entries.Count)
            : 304_800L;

        var gf = BuildTableFrame(
            indexPart, entries, wordPartMap, columns, colWidths,
            x, y, cx, hdrH, rowH, rowOffset, hyperlink);

        slide.CommonSlideData?.ShapeTree?.Append(gf);
        indexPart.Slide.Save();
    }

    private static P.Shape? FindShapeContainingText(P.Slide slide, string text) =>
        slide.Descendants<P.Shape>()
             .FirstOrDefault(sp => sp.InnerText.Contains(text));

    // =======================================================================
    // GraphicFrame + Table
    // =======================================================================

    private static P.GraphicFrame BuildTableFrame(
        SlidePart                        indexPart,
        List<WordEntry>                  entries,
        Dictionary<WordEntry, SlidePart> wordPartMap,
        List<string>                     columns,
        List<long>                       colWidths,
        long x, long y, long cx,
        long hdrH, long rowH,
        int  rowOffset,
        bool hyperlink)
    {
        long frameCy = hdrH + rowH * entries.Count;

        var gf = new P.GraphicFrame();

        gf.Append(new P.NonVisualGraphicFrameProperties(
            new P.NonVisualDrawingProperties { Id = 10, Name = "Index Table" },
            new P.NonVisualGraphicFrameDrawingProperties(
                new A.GraphicFrameLocks { NoGrouping = true }),
            new ApplicationNonVisualDrawingProperties()));

        gf.Append(new P.Transform(
            new A.Offset  { X = x,  Y = y  },
            new A.Extents { Cx = cx, Cy = frameCy }));

        var tbl   = new A.Table();
        var tblPr = new A.TableProperties { FirstRow = true, BandRow = true };
        tblPr.Append(new A.TableStyleId
            { Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}" });
        tbl.Append(tblPr);

        var tblGrid = new A.TableGrid();
        foreach (var w in colWidths)
            tblGrid.Append(new A.GridColumn { Width = w });
        tbl.Append(tblGrid);

        tbl.Append(BuildHeaderRow(columns, hdrH));

        for (int i = 0; i < entries.Count; i++)
        {
            tbl.Append(BuildDataRow(
                indexPart, entries[i], wordPartMap, columns,
                rowNumber: i + rowOffset + 1,
                height:    rowH,
                altRow:    i % 2 == 1,
                hyperlink: hyperlink));
        }

        gf.Append(new A.Graphic(
            new A.GraphicData(tbl)
            {
                Uri = "http://schemas.openxmlformats.org/drawingml/2006/table"
            }));

        return gf;
    }

    // ── Header row ────────────────────────────────────────────────────────

    private static A.TableRow BuildHeaderRow(List<string> columns, long height)
    {
        var tr = new A.TableRow { Height = height };

        foreach (var col in columns)
        {
            var tc     = new A.TableCell();
            var txBody = new A.TextBody(
                new A.BodyProperties(),
                new A.ListStyle(),
                MakeParagraph(ColDefs[col].Header,
                              bold: true, fontSize: 12,
                              colorHex: "FFFFFF",
                              align: A.TextAlignmentTypeValues.Left));
            tc.Append(txBody);

            var tcPr = new A.TableCellProperties();
            tcPr.Append(new A.SolidFill(
                new A.RgbColorModelHex { Val = "1C4E80" }));
            tc.Append(tcPr);
            tr.Append(tc);
        }

        return tr;
    }

    // ── Data row ──────────────────────────────────────────────────────────

    private static A.TableRow BuildDataRow(
        SlidePart                        indexPart,
        WordEntry                        entry,
        Dictionary<WordEntry, SlidePart> wordPartMap,
        List<string>                     columns,
        int                              rowNumber,
        long                             height,
        bool                             altRow,
        bool                             hyperlink)
    {
        var tr         = new A.TableRow { Height = height };
        int wordColIdx = columns.IndexOf("{word}");

        for (int c = 0; c < columns.Count; c++)
        {
            string token  = columns[c];
            string value  = GetEntryValue(entry, token, rowNumber);
            bool   isWord = c == wordColIdx;

            var         tc   = new A.TableCell();
            A.Paragraph para;

            if (isWord && hyperlink && wordPartMap.TryGetValue(entry, out SlidePart? targetPart))
            {
                string relId = AddSlideJumpRelationship(indexPart, targetPart);
                para = MakeHyperlinkParagraph(value, relId, fontSize: 11);
            }
            else
            {
                para = MakeParagraph(value,
                    bold:     false,
                    fontSize: 11,
                    colorHex: "333333",
                    align:    token == "{n}"
                              ? A.TextAlignmentTypeValues.Center
                              : A.TextAlignmentTypeValues.Left);
            }

            tc.Append(new A.TextBody(
                new A.BodyProperties(),
                new A.ListStyle(),
                para));

            var tcPr = new A.TableCellProperties();
            if (altRow)
                tcPr.Append(new A.SolidFill(
                    new A.RgbColorModelHex { Val = "F2F2F2" }));
            tc.Append(tcPr);
            tr.Append(tc);
        }

        return tr;
    }

    // ── Paragraph builders ────────────────────────────────────────────────

    private static A.Paragraph MakeParagraph(
        string text, bool bold, int fontSize, string colorHex,
        A.TextAlignmentTypeValues align)
    {
        var pPr = new A.ParagraphProperties { Alignment = align };

        var rPr = new A.RunProperties { Language = "en-US", Dirty = false };
        rPr.Append(new A.SolidFill(new A.RgbColorModelHex { Val = colorHex }));
        if (bold) rPr.Bold = true;
        rPr.FontSize = fontSize * 100;

        var run = new A.Run();
        run.Append(rPr);
        run.Append(new A.Text(text));

        var para = new A.Paragraph();
        para.Append(pPr);
        para.Append(run);
        return para;
    }

    private static A.Paragraph MakeHyperlinkParagraph(
        string text, string relId, int fontSize)
    {
        var pPr = new A.ParagraphProperties
            { Alignment = A.TextAlignmentTypeValues.Left };

        var rPr = new A.RunProperties { Language = "en-US", Dirty = false };
        rPr.Append(new A.SolidFill(new A.RgbColorModelHex { Val = "0073AE" }));
        rPr.FontSize = fontSize * 100;

        // action="ppaction://hlinksldjump" is what makes PowerPoint jump slides.
        // Without it the link is ignored, even with a valid relationship.
        rPr.InsertAt(
            new A.HyperlinkOnClick { Id = relId, Action = SlideJumpAction }, 0);

        var run = new A.Run();
        run.Append(rPr);
        run.Append(new A.Text(text));

        var para = new A.Paragraph();
        para.Append(pPr);
        para.Append(run);
        return para;
    }

    // ── Value getter ──────────────────────────────────────────────────────

    private static string GetEntryValue(WordEntry e, string token, int n) => token switch
    {
        "{n}"       => n.ToString(),
        "{word}"    => e.Word,
        "{plural}"  => e.Plural,
        "{type}"    => e.Type,
        "{english}" => e.English,
        _           => "",
    };

    // =======================================================================
    // Word slide processing
    // =======================================================================

    private static void ProcessWordSlide(
        SlidePart      slidePart,
        WordEntry      entry,
        TemplateConfig config)
    {
        MergeTextRuns(slidePart);
        SubstitutePlaceholders(slidePart, entry.ToPlaceholders());

        if (entry.HasImage)
            InjectImage(slidePart, entry.Image!, ShapeImage);
        else
            RemoveShapeByName(slidePart, ShapeImage);

        if (entry.HasAudio)
            InjectAudio(slidePart, entry.Audio!, ShapeAudio);
        else
            RemoveShapeByName(slidePart, ShapeAudio);

        slidePart.Slide.Save();
    }

    // =======================================================================
    // Navigation shapes (shape_next, shape_index on word slides)
    // =======================================================================

    private static void WireNavigationShapes(
        SlidePart  wordPart,
        SlidePart? nextPart,
        SlidePart? indexPart)
    {
        var nextShape = FindShapeByName(wordPart, ShapeNext);
        if (nextShape != null && nextPart != null)
        {
            string relId = AddSlideJumpRelationship(wordPart, nextPart);
            SetShapeHlinkClick(nextShape, relId);
        }

        var idxShape = FindShapeByName(wordPart, ShapeIndex);
        if (idxShape != null && indexPart != null)
        {
            string relId = AddSlideJumpRelationship(wordPart, indexPart);
            SetShapeHlinkClick(idxShape, relId);
        }

        wordPart.Slide.Save();
    }

    private static void SetShapeHlinkClick(P.Shape shape, string relId)
    {
        var cNvPr = shape.NonVisualShapeProperties?.NonVisualDrawingProperties;
        if (cNvPr == null) return;

        cNvPr.Elements<A.HyperlinkOnClick>().ToList().ForEach(h => h.Remove());

        foreach (var rPr in shape.Descendants<A.RunProperties>())
            rPr.Elements<A.HyperlinkOnClick>().ToList().ForEach(h => h.Remove());

        cNvPr.InsertAt(new A.HyperlinkOnClick
        {
            Id = relId,
            Action = SlideJumpAction
        }, 0);
    }

    // =======================================================================
    // Shape finders
    // =======================================================================

    private static P.Shape? FindShapeByName(SlidePart slidePart, string name) =>
        slidePart.Slide.Descendants<P.Shape>()
            .FirstOrDefault(sp => string.Equals(
                sp.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value,
                name, StringComparison.OrdinalIgnoreCase));

    private static void RemoveShapeByName(SlidePart slidePart, string name) =>
        FindShapeByName(slidePart, name)?.Remove();

    // =======================================================================
    // Image injection
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
        using (var stream = asset.OpenStream()) imagePart.FeedData(stream);

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
    // Audio injection
    // =======================================================================

    private static void InjectAudio(SlidePart slidePart, AssetData asset, string shapeName)
    {
        var shape = FindShapeByName(slidePart, shapeName);

        long x = 457_200, y = 457_200, cx = 457_200, cy = 457_200;
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
            mediaPart = slidePart.OpenXmlPackage.CreateMediaDataPart(
                "audio/mpeg", ".mp3");
        }

        using (var s = asset.OpenStream()) mediaPart.FeedData(s);

        var audioRelId = slidePart.AddAudioReferenceRelationship(mediaPart).Id;
        var mediaRelId = slidePart.AddMediaReferenceRelationship(mediaPart).Id;

        var audioShape = BuildAudioShape(audioRelId, mediaRelId, shapeName, x, y, cx, cy);
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
        cNvPr.Append(new A.HyperlinkOnClick
            { Id = audioRelId, Action = "ppaction://media" });
        nvPPr.Append(cNvPr);
        nvPPr.Append(new P.NonVisualPictureDrawingProperties());
        var appNvPr = new ApplicationNonVisualDrawingProperties();
        appNvPr.Append(new AudioFromFile { Link = mediaRelId });
        nvPPr.Append(appNvPr);
        pic.Append(nvPPr);
        pic.Append(new P.BlipFill(
            new A.Blip(), new A.Stretch(new A.FillRectangle())));
        pic.Append(new P.ShapeProperties(
            new A.Transform2D(
                new A.Offset  { X = x, Y = y },
                new A.Extents { Cx = cx, Cy = cy }),
            new A.PresetGeometry(new A.AdjustValueList())
                { Preset = A.ShapeTypeValues.Rectangle }));
        return pic;
    }

    // =======================================================================
    // Helpers
    // =======================================================================

    private static void SetSlideName(SlidePart slidePart, string name)
    {
        var cSld = slidePart.Slide.CommonSlideData;
        if (cSld != null) cSld.Name = name;
    }

    private static Uri BuildSlideJumpUri(SlidePart slidePart)
    {
        var slideName = slidePart.Slide.CommonSlideData?.Name?.Value;
        if (!string.IsNullOrWhiteSpace(slideName))
            return new Uri($"#{slideName}", UriKind.Relative);

        // Fallback for unexpected unnamed slides.
        return slidePart.Uri;
    }

    private static string SanitizeName(string word)
    {
        var sb = new StringBuilder();
        foreach (char c in word)
            sb.Append(char.IsLetterOrDigit(c) ? c : '_');
        return sb.ToString().Trim('_');
    }

    private static SlidePart CloneSlide(PresentationPart presPart, SlidePart source)
    {
        var newPart = presPart.AddNewPart<SlidePart>();
        using (var stream = source.GetStream()) newPart.FeedData(stream);
        foreach (var rel in source.Parts)
            newPart.AddPart(rel.OpenXmlPart, rel.RelationshipId);
        return newPart;
    }

    private static void SubstitutePlaceholders(
        SlidePart slidePart, Dictionary<string, string> map)
    {
        string xml;
        using (var r = new StreamReader(slidePart.GetStream())) xml = r.ReadToEnd();
        foreach (var (key, value) in map)
            xml = xml.Replace(key, System.Security.SecurityElement.Escape(value ?? ""));
        using var w = new StreamWriter(slidePart.GetStream(FileMode.Create));
        w.Write(xml);
    }

    private static void MergeTextRuns(SlidePart slidePart)
    {
        foreach (var para in slidePart.Slide.Descendants<A.Paragraph>().ToList())
        {
            var runs = para.Elements<A.Run>().ToList();
            if (runs.Count < 2) continue;
            var combined = string.Concat(runs.Select(r => r.Text?.Text ?? ""));
            if (!combined.Contains("{{")) continue;
            runs[0].Text = new A.Text(combined);
            for (int i = 1; i < runs.Count; i++) runs[i].Remove();
        }
    }

    // =======================================================================
    // Internal data carrier
    // =======================================================================

    private class OutputSlide
    {
        public SlidePart       Part            { get; }
        public SlideRole       Role            { get; }
        public WordEntry?      Entry           { get; init; }
        public List<WordEntry> IndexEntries    { get; init; } = [];
        public int             IndexPage       { get; init; } = 1;
        public int             GlobalRowOffset { get; init; } = 0;

        public OutputSlide(SlidePart part, SlideRole role)
        {
            Part = part;
            Role = role;
        }
    }
}
