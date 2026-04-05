using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using SlideshowGenerator.Models;
using System.IO.Packaging;
using System.Xml;

namespace SlideshowGenerator.Services;

public static class ExcelReader
{
    private static readonly string[] AudioExtensions =
        [".mp3", ".wav", ".m4a", ".ogg"];

    // -----------------------------------------------------------------------
    // Public entry point
    // -----------------------------------------------------------------------

    public static List<WordEntry> Read(string excelPath)
    {
        var entries = new List<WordEntry>();
        var baseDir = Path.GetDirectoryName(excelPath) ?? "";

        using var wb = new XLWorkbook(excelPath);
        var ws = wb.Worksheets.First();

        // Pre-extract all embedded assets keyed by (row, col) — 1-based
        var embeddedImages = ExtractEmbeddedImages(ws);
        var embeddedAudio  = ExtractEmbeddedOleObjects(ws, excelPath);

        foreach (var row in ws.RowsUsed().Skip(1)) // row 1 = header
        {
            var word = row.Cell(1).GetString().Trim();
            if (string.IsNullOrWhiteSpace(word)) continue;

            int rowNum = row.RowNumber();

            entries.Add(new WordEntry
            {
                Word          = word,
                Plural        = row.Cell(2).GetString().Trim(),
                Stress        = row.Cell(3).GetString().Trim(),
                Vowels        = row.Cell(4).GetString().Trim(),
                Hint          = row.Cell(5).GetString().Trim(),
                Trap          = row.Cell(6).GetString().Trim(),
                Type          = row.Cell(7).GetString().Trim(),
                Rule          = row.Cell(8).GetString().Trim(),
                Usage         = row.Cell(9).GetString().Trim(),
                English       = row.Cell(10).GetString().Trim(),
                Pronunciation = row.Cell(11).GetString().Trim(),
                Image = ResolveImage(row.Cell(12), rowNum, baseDir, embeddedImages),
                Audio = ResolveAudio(row.Cell(13), rowNum, baseDir, embeddedAudio),
            });
        }

        return entries;
    }

    // -----------------------------------------------------------------------
    // Asset resolution — path wins over embedded if both somehow exist
    // -----------------------------------------------------------------------

    private static AssetData? ResolveImage(
        IXLCell cell, int row, string baseDir,
        Dictionary<(int r, int c), (byte[] bytes, string ext)> embedded)
    {
        var text = cell.GetString().Trim();
        if (!string.IsNullOrWhiteSpace(text))
        {
            var asset = AssetData.FromPath(text, baseDir);
            if (asset != null) return asset;
        }

        if (embedded.TryGetValue((row, 12), out var pic))
            return AssetData.FromBytes(pic.bytes, pic.ext);

        return null;
    }

    private static AssetData? ResolveAudio(
        IXLCell cell, int row, string baseDir,
        Dictionary<(int r, int c), (byte[] bytes, string ext)> embedded)
    {
        var text = cell.GetString().Trim();
        if (!string.IsNullOrWhiteSpace(text))
        {
            var asset = AssetData.FromPath(text, baseDir);
            if (asset != null) return asset;
        }

        if (embedded.TryGetValue((row, 13), out var ole))
            return AssetData.FromBytes(ole.bytes, ole.ext);

        return null;
    }

    // -----------------------------------------------------------------------
    // Extract pictures placed in cells via Insert → Pictures → Place in Cell
    // -----------------------------------------------------------------------

    private static Dictionary<(int r, int c), (byte[] bytes, string ext)>
        ExtractEmbeddedImages(IXLWorksheet ws)
    {
        var result = new Dictionary<(int, int), (byte[], string)>();

        foreach (var picture in ws.Pictures)
        {
            int row = picture.TopLeftCell.Address.RowNumber;
            int col = picture.TopLeftCell.Address.ColumnNumber;
            if (col != 12) continue;

            using var ms = new MemoryStream();
            picture.ImageStream.CopyTo(ms);

            var ext = picture.Format switch
            {
                XLPictureFormat.Jpeg => ".jpg",
                XLPictureFormat.Png  => ".png",
                XLPictureFormat.Gif  => ".gif",
                XLPictureFormat.Bmp  => ".bmp",
                _                    => ".png"
            };

            result[(row, col)] = (ms.ToArray(), ext);
        }

        return result;
    }

    // -----------------------------------------------------------------------
    // Extract OLE objects (Insert → Object) — used for audio files
    // ClosedXML doesn't expose these; we read the raw package directly.
    // -----------------------------------------------------------------------

    private static Dictionary<(int r, int c), (byte[] bytes, string ext)>
        ExtractEmbeddedOleObjects(IXLWorksheet ws, string excelPath)
    {
        var result = new Dictionary<(int, int), (byte[], string)>();

        try
        {
            using var package = Package.Open(excelPath, FileMode.Open, FileAccess.Read);

            var wsUri = GetWorksheetPartUri(package, ws.Name);
            if (wsUri == null) return result;

            var wsPart = package.GetPart(wsUri);
            var wsXml  = LoadXml(wsPart);
            var ns     = BuildNsManager(wsXml);

            var oleNodes = wsXml.SelectNodes("//x:oleObjects/x:oleObject", ns);
            if (oleNodes == null) return result;

            foreach (XmlNode oleNode in oleNodes)
            {
                var rId     = oleNode.Attributes?["r:id"]?.Value;
                var shapeId = oleNode.Attributes?["shapeId"]?.Value;
                if (rId == null) continue;

                var rel = wsPart.GetRelationship(rId);
                if (rel == null) continue;

                var targetUri = PackUriHelper.ResolvePartUri(wsUri, rel.TargetUri);
                if (!package.PartExists(targetUri)) continue;

                var bytes = ReadPartBytes(package.GetPart(targetUri));
                var ext   = GuessAudioExtension(bytes, rel.TargetUri.ToString());

                var anchor = FindDrawingAnchor(package, wsUri, wsPart, shapeId);
                if (anchor == null) continue;

                // Force column 13 — OLE objects in col 13 are treated as audio
                if (anchor.Value.col == 13)
                    result[(anchor.Value.row, 13)] = (bytes, ext);
            }
        }
        catch
        {
            // Non-fatal — if OLE extraction fails, embedded audio is simply unavailable
        }

        return result;
    }

    // -----------------------------------------------------------------------
    // Package XML helpers
    // -----------------------------------------------------------------------

    private static Uri? GetWorksheetPartUri(Package package, string sheetName)
    {
        var wbUri  = new Uri("/xl/workbook.xml", UriKind.Absolute);
        var wbPart = package.GetPart(wbUri);
        var wbXml  = LoadXml(wbPart);
        var ns     = BuildNsManager(wbXml);

        var node = wbXml.SelectSingleNode($"//x:sheet[@name='{sheetName}']", ns);
        var rId  = node?.Attributes?["r:id"]?.Value;
        if (rId == null) return null;

        var rel = wbPart.GetRelationship(rId);
        return PackUriHelper.ResolvePartUri(wbUri, rel.TargetUri);
    }

    private static (int row, int col)? FindDrawingAnchor(
        Package package, Uri wsUri, PackagePart wsPart, string? shapeId)
    {
        if (shapeId == null) return null;

        var drawingRel = wsPart.GetRelationships()
            .FirstOrDefault(r => r.RelationshipType.EndsWith("/drawing"));
        if (drawingRel == null) return null;

        var drawingUri = PackUriHelper.ResolvePartUri(wsUri, drawingRel.TargetUri);
        if (!package.PartExists(drawingUri)) return null;

        var xml = LoadXml(package.GetPart(drawingUri));
        var ns  = BuildNsManager(xml);

        var anchors = xml.SelectNodes("//xdr:twoCellAnchor", ns);
        if (anchors == null) return null;

        foreach (XmlNode anchor in anchors)
        {
            var cNvPr = anchor.SelectSingleNode(".//xdr:sp/xdr:nvSpPr/xdr:cNvPr", ns);
            if (cNvPr?.Attributes?["id"]?.Value != shapeId) continue;

            var from = anchor.SelectSingleNode("xdr:from", ns);
            var colTxt = from?.SelectSingleNode("xdr:col", ns)?.InnerText;
            var rowTxt = from?.SelectSingleNode("xdr:row", ns)?.InnerText;

            if (colTxt == null || rowTxt == null) continue;

            // xdr values are 0-based
            return (int.Parse(rowTxt) + 1, int.Parse(colTxt) + 1);
        }

        return null;
    }

    private static string GuessAudioExtension(byte[] bytes, string uriHint)
    {
        var ext = Path.GetExtension(uriHint).ToLowerInvariant();
        if (AudioExtensions.Contains(ext)) return ext;

        // Magic-byte fallback
        if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFB) return ".mp3";
        if (bytes.Length >= 4 &&
            bytes[0] == 'R' && bytes[1] == 'I' && bytes[2] == 'F' && bytes[3] == 'F') return ".wav";

        return ".mp3";
    }

    private static XmlDocument LoadXml(PackagePart part)
    {
        var doc = new XmlDocument();
        using var s = part.GetStream();
        doc.Load(s);
        return doc;
    }

    private static XmlNamespaceManager BuildNsManager(XmlDocument doc)
    {
        var ns = new XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("x",   "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        ns.AddNamespace("r",   "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        ns.AddNamespace("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
        return ns;
    }

    private static byte[] ReadPartBytes(PackagePart part)
    {
        using var ms = new MemoryStream();
        using var s  = part.GetStream();
        s.CopyTo(ms);
        return ms.ToArray();
    }
}
