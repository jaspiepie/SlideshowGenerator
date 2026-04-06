namespace LanguageCourseSlides.Forms;

public class HelpForm : Form
{
    public HelpForm()
    {
        BuildUI();
    }

    private void BuildUI()
    {
        Text            = "Language Course Slide Generator — Help & Instructions";
        Size            = new Size(820, 640);
        MinimumSize     = new Size(640, 480);
        StartPosition   = FormStartPosition.CenterParent;
        Font            = new Font("Segoe UI", 9f);

        var tabs = new TabControl { Dock = DockStyle.Fill };

        tabs.TabPages.Add(MakeTab("Overview",        OverviewText()));
        tabs.TabPages.Add(MakeTab("Excel Format",    ExcelFormatText()));
        tabs.TabPages.Add(MakeTab("Template Shapes", ShapesText()));
        tabs.TabPages.Add(MakeTab("Template Setup",  TemplateSetupText()));
        tabs.TabPages.Add(MakeTab("Index & Linking", IndexText()));

        var btnClose = new Button
        {
            Text        = "Close",
            Dock        = DockStyle.Bottom,
            Height      = 32,
            DialogResult = DialogResult.OK,
        };

        Controls.Add(tabs);
        Controls.Add(btnClose);
        AcceptButton = btnClose;
        CancelButton = btnClose;
    }

    private static TabPage MakeTab(string title, string content)
    {
        var rtb = new RichTextBox
        {
            Dock      = DockStyle.Fill,
            ReadOnly  = true,
            BackColor = SystemColors.Window,
            BorderStyle = BorderStyle.None,
            Font      = new Font("Segoe UI", 9.5f),
            ScrollBars = RichTextBoxScrollBars.Vertical,
        };

        // Parse simple markdown-like: lines starting with ## are headers
        foreach (var line in content.Split('\n'))
        {
            if (line.StartsWith("## "))
            {
                rtb.SelectionFont  = new Font("Segoe UI", 11f, FontStyle.Bold);
                rtb.SelectionColor = Color.FromArgb(28, 78, 128);
                rtb.AppendText(line[3..] + "\n");
                rtb.SelectionFont  = new Font("Segoe UI", 9.5f);
                rtb.SelectionColor = Color.Black;
            }
            else if (line.StartsWith("  ") && line.TrimStart().StartsWith("•"))
            {
                rtb.SelectionFont        = new Font("Segoe UI", 9.5f);
                rtb.SelectionBullet      = false;
                rtb.SelectionIndent      = 20;
                rtb.SelectionHangingIndent = 0;
                rtb.AppendText(line.TrimStart() + "\n");
                rtb.SelectionIndent = 0;
            }
            else
            {
                rtb.SelectionFont  = new Font("Segoe UI", 9.5f);
                rtb.SelectionColor = Color.Black;
                rtb.AppendText(line + "\n");
            }
        }
        rtb.SelectionStart = 0;
        rtb.ScrollToCaret();

        var page = new TabPage(title);
        page.Controls.Add(rtb);
        page.Padding = new Padding(8);
        return page;
    }

    // ── Tab content ─────────────────────────────────────────────────────────

    private static string OverviewText() => @"
## What This App Does

Language Course Slide Generator turns an Excel word list into a fully-formatted PowerPoint presentation. It reads each row of your spreadsheet and clones a word slide template once per word, filling in all the linguistic detail automatically.

## Workflow

  • Step 1 — Create or select a Template Configuration (Manage Templates button).
  • Step 2 — Load your word list Excel file (Browse button).
  • Step 3 — Review the preview grid. Click any Image or Audio cell to assign a file.
  • Step 4 — Click Generate Presentation and choose where to save.

## Output

The generated file contains:
  • All Static slides (covers, intro pages) copied once.
  • One or more Index slides listing every word with clickable hyperlinks.
  • One Word slide per entry, fully populated with text, image, and audio.

Each slide is named — e.g. Slide_Index, Slide_blau, Slide_der_Hund — so you can navigate by name in PowerPoint's slide panel.

## Template Configuration

A Template Configuration tells the app which slides in your .pptx file play which role (Static, Index, or Word). You can save as many configurations as you like and switch between them.
".TrimStart();

    private static string ExcelFormatText() => @"
## Excel Column Layout

Row 1 is always the header row and is skipped. Data starts at row 2.

  • Col 1   Word                  Required. The German word as it appears on the slide.
  • Col 2   Plural                Optional. Plural form (e.g. die Hunde).
  • Col 3   Stress                Stressed syllable(s), e.g. BLAU or KAT-ze.
  • Col 4   Vowels                Vowel description, e.g. diphthong au.
  • Col 5   Hint                  IPA or memory hook.
  • Col 6   Trap                  Common learner mistake.
  • Col 7   Type                  Word type: Adjective, Noun, Verb, Adverb, etc.
  • Col 8   Rule                  Optional grammar rule shown on the slide.
  • Col 9   Usage                 Optional example sentence.
  • Col 10  English               English translation — shown prominently top-right.
  • Col 11  Pronunciation         Optional detailed pronunciation guide.
  • Col 12  Image                 Optional. File path or embedded picture.
  • Col 13  Audio                 Optional. File path or embedded audio object.

## Image Column (12)

Three options:
  • Leave blank — no image on this word's slide (shape_image is removed).
  • Type or paste a file path — absolute (C:\files\hund.jpg) or relative to the Excel file.
  • Embed directly — Insert → Pictures → Place in Cell on that cell, leave text empty.

## Audio Column (13)

Three options:
  • Leave blank — audio shape is removed from the slide.
  • Type or paste a file path (.mp3, .wav, .m4a, .ogg).
  • Embed directly — Insert → Object, browse to your audio file, leave text empty.

You can also click any Image or Audio cell in the app's preview grid to browse for a file after loading.

## Saving an Excel Template

Use File → Save Excel Template to download a blank spreadsheet with all column headers already set up and an Instructions tab explaining each column.
".TrimStart();

    private static string ShapesText() => @"
## Shape Name Conventions

The generator finds special shapes in your PowerPoint template by their name. To name a shape in PowerPoint, right-click it → Edit Alt Text, or use the Selection Pane (Home → Arrange → Selection Pane) and double-click the name.

## Word Slide Shapes

Name these shapes in your Word slide template:

  • shape_audio    The speaker / listen button. If the word has audio, this shape is made into a clickable audio trigger. If there is no audio for this word, the shape is removed entirely from that slide.

  • shape_image    The image placeholder. If the word has an image, this shape is replaced by the actual picture at the same position and size. If there is no image, the shape is removed.

  • shape_next     Navigation button. Automatically hyperlinked to the next word slide. On the last word slide, links back to the index.

  • shape_index    Back-to-index button. Automatically hyperlinked to the first index slide.

## Index Slide

On your Index slide template, place a text box containing exactly:

    {{Index}}

The app finds this text box and replaces it with one line per word. The text box's font, size, and colour are inherited by the generated list. If hyperlinks are enabled, each word is a clickable link to its slide.

## Text Placeholders (Word Slide)

Any text box containing one of these strings will have it replaced:

  • {{Word}}           The German word
  • {{Plural}}         Plural form
  • {{Stress}}         Stress pattern
  • {{Vowels}}         Vowel notes
  • {{Hint}}           IPA / memory hint
  • {{Trap}}           Common mistake
  • {{Type}}           Word type
  • {{Rule}}           Grammar rule
  • {{Usage}}          Example usage
  • {{English}}        English translation
  • {{Pronunciation}}  Pronunciation guide

If a field is empty, the placeholder is replaced with an empty string — the text box remains but is blank.
".TrimStart();

    private static string TemplateSetupText() => @"
## Creating a Template Configuration

Click Manage Templates → New Template to open the editor.

## Fields

  • Template name     A display name for this configuration (e.g. German A1 Standard).
  • Template file     Browse to your .pptx file. The app reads the slide count automatically.
  • Words per page    How many index entries fit on one index slide before a new page is created. Default: 20.
  • Index format      How each line in the index is formatted. Available tokens:
                        {word}    the German word
                        {plural}  plural form
                        {type}    word type
                        {english} English translation
                        {n}       line number
  • Hyperlink index   When on, each index entry is a clickable hyperlink to its word slide.

## Slide Role Assignment

After loading a .pptx, each slide is shown in a grid. Assign a role to each:

  • Static   Copied once, unchanged. Use for cover slides, intro pages, background art.
  • Index    Copied once per batch (if word count exceeds Words per Page, extra copies are made).
             The {{Index}} text box is replaced with the word list.
  • Word     Cloned once per word entry. All {{Placeholder}} text and named shapes are processed.

Rules:
  • Exactly one slide must have the Word role.
  • At most one slide template can have the Index role (it is cloned if overflow is needed).
  • Any number of slides can be Static.

## Editing an Existing Configuration

Open Manage Templates, select a configuration, click Edit. Changes take effect on the next generation run.
".TrimStart();

    private static string IndexText() => @"
## How the Index Works

After all word slides are built, the app goes back and fills in the Index slide(s).

## Index Line Format → Table Columns

The Index Line Format field controls which columns appear in the index table and in what order. Each `{token}` you include becomes one column. The literal text between tokens is ignored — only the tokens matter.

Examples:
  • {word}                         → 1 column:  Word
  • {word} {type} {english}        → 3 columns: Word, Type, English
  • {n} {word} {plural} {type} {english}  → 5 columns: #, Word, Plural, Type, English

## Hyperlinks

When Hyperlink Index is on, each word in the index is a clickable internal hyperlink to the corresponding word slide. In presentation mode, clicking a word jumps directly to that slide.

## Index Overflow

If you have more words than the Words Per Page setting allows, the app clones the index slide template automatically — one copy per batch. For example, with 45 words and Words Per Page = 20, you get three index slides (20 / 20 / 5 words each).

All index slide copies are inserted in sequence before the word slides.

## shape_index Navigation

Every word slide can have a shape named shape_index. The app automatically wires this shape as a hyperlink back to the first index slide. This creates a back-navigation button without any manual linking.

## shape_next Navigation

Every word slide can have a shape named shape_next. The app wires it to the next slide in sequence. On the last word slide, shape_next links back to the index.

## Slide Naming

Every output slide is given a descriptive name visible in PowerPoint's slide panel:
  • Slide_Index (or Slide_Index_1, Slide_Index_2 if multiple pages)
  • Slide_blau, Slide_braun, Slide_der_Hund, etc.
  • Slide_001_static for static slides
".TrimStart();
}
