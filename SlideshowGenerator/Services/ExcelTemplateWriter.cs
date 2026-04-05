using ClosedXML.Excel;

namespace LanguageCourseSlides.Services;

public static class ExcelTemplateWriter
{
    public static void Save(string path)
    {
        using var wb = new XLWorkbook();

        // ── Word List sheet ──────────────────────────────────────────────
        var ws = wb.AddWorksheet("Word List");

        string[] headers =
        [
            "Word", "Plural", "Stress", "Vowels", "Hint",
            "Trap", "Type", "Rule", "Usage", "English",
            "Pronunciation Explained", "Image", "Audio"
        ];

        // Header row
        for (int col = 1; col <= headers.Length; col++)
        {
            var cell = ws.Cell(1, col);
            cell.Value = headers[col - 1];
            cell.Style.Font.Bold      = true;
            cell.Style.Font.FontColor = XLColor.White;
            cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#1C4E80");
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        }

        // Example row 2 — blau
        object[] example =
        [
            "blau", "", "BLAU", "diphthong au", "/blaʊ̯/",
            "None.", "Adjective", "Primary color.", "", "blue",
            "One syllable. 'au' sounds like 'ow' in cow.", "", ""
        ];
        for (int col = 1; col <= example.Length; col++)
            ws.Cell(2, col).Value = (string)example[col - 1];

        // Column widths
        int[] widths = [ 18, 14, 12, 22, 14, 24, 14, 24, 28, 14, 40, 14, 14 ];
        for (int col = 1; col <= widths.Length; col++)
            ws.Column(col).Width = widths[col - 1];

        ws.Row(1).Height = 22;
        ws.SheetView.FreezeRows(1);
        ws.RangeUsed()!.SetAutoFilter();

        // ── Instructions sheet ───────────────────────────────────────────
        var wi = wb.AddWorksheet("Instructions");
        wi.ShowGridLines = false;

        var instructions = new (string text, bool bold, int size, string colour)[]
        {
            ("Language Course Slide Generator — Word List Format", true,  13, "1C4E80"),
            ("",                                                   false, 10, "000000"),
            ("Column Reference",                                   true,  11, "1C4E80"),
            ("Col 1   Word                 The German word as shown on the slide", false, 10, "333333"),
            ("Col 2   Plural               Optional plural form",                  false, 10, "333333"),
            ("Col 3   Stress               Stressed syllable(s), e.g. KAT-ze",     false, 10, "333333"),
            ("Col 4   Vowels               Vowel sounds and notes",                 false, 10, "333333"),
            ("Col 5   Hint                 IPA or memory hook",                     false, 10, "333333"),
            ("Col 6   Trap                 Common learner mistake",                 false, 10, "333333"),
            ("Col 7   Type                 Adjective / Noun / Verb / Adverb",       false, 10, "333333"),
            ("Col 8   Rule                 Optional grammar rule",                  false, 10, "333333"),
            ("Col 9   Usage                Optional example sentence",              false, 10, "333333"),
            ("Col 10  English              English translation",                    false, 10, "333333"),
            ("Col 11  Pronunciation        Detailed pronunciation guide",           false, 10, "333333"),
            ("",                                                   false, 10, "000000"),
            ("Asset Columns",                                      true,  11, "1C4E80"),
            ("Col 12  Image    Leave blank = no image (shape_image removed from slide).", false, 10, "333333"),
            ("         Type a file path, or embed via Insert → Pictures → Place in Cell.", false, 10, "555555"),
            ("Col 13  Audio    Leave blank = no audio (shape_audio removed from slide).",  false, 10, "333333"),
            ("         Type a file path (.mp3/.wav), or embed via Insert → Object.",        false, 10, "555555"),
            ("",                                                   false, 10, "000000"),
            ("Shape Name Conventions (in PowerPoint template)",    true,  11, "1C4E80"),
            ("shape_audio    Speaker/listen button. Made into audio trigger, or removed.", false, 10, "333333"),
            ("shape_image    Image placeholder. Replaced by picture, or removed.",         false, 10, "333333"),
            ("shape_next     Next-slide button. Auto-hyperlinked to next word slide.",     false, 10, "333333"),
            ("shape_index    Back-to-index button. Auto-hyperlinked to index slide.",      false, 10, "333333"),
            ("{{Index}}      Text box on index slide — replaced with word list.",          false, 10, "333333"),
            ("",                                                   false, 10, "000000"),
            ("Row 1 is always the header row and is skipped on import.", false, 9, "777777"),
        };

        for (int r = 1; r <= instructions.Length; r++)
        {
            var (text, bold, size, colour) = instructions[r - 1];
            var cell = wi.Cell(r, 1);
            cell.Value = text;
            cell.Style.Font.Bold      = bold;
            cell.Style.Font.FontSize  = size;
            cell.Style.Font.FontColor = XLColor.FromHtml($"#{colour}");
        }
        wi.Column(1).Width = 80;

        wb.SaveAs(path);
    }
}
