# Language Course Slide Generator

A WinForms application that generates PowerPoint presentations from Excel word lists,
using configurable template profiles with support for image and audio assets,
index slides, and internal hyperlinks.

---

## Requirements

- Windows 10 or later
- [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0) to build
- Visual Studio 2022+ recommended

---

## Building

```
cd SlideshowGenerator
dotnet restore
dotnet build
dotnet run
```

Or open `SlideshowGenerator.csproj` in Visual Studio and press F5.

---

## Project Structure

```
SlideshowGenerator/
├── Models/
│   ├── AssetData.cs           Carries image/audio from a file path or embedded bytes
│   ├── SlideDefinition.cs     Maps a template slide to a role (Static / Index / Word)
│   ├── TemplateConfig.cs      Full saved template configuration
│   └── WordEntry.cs           One row from the Excel word list
├── Services/
│   ├── ConfigManager.cs       Saves/loads template configs as JSON (AppData)
│   ├── ExcelReader.cs         Reads word list + resolves file-path and embedded assets
│   └── PptxGenerator.cs       Two-pass PPTX builder with hyperlinks, images, audio
├── Forms/
│   ├── MainForm.cs            Main application window
│   ├── TemplateConfigForm.cs  Create / edit a template configuration
│   └── TemplateManagerForm.cs List, edit, and delete saved templates
└── Program.cs
```

---

## Slide Roles

| Role   | Behaviour |
|--------|-----------|
| Static | Copied once, unchanged (covers, intro images, etc.) |
| Index  | Copied once; `{{Index}}` replaced with the word list |
| Word   | Cloned once **per word entry**; all placeholders substituted |

---

## Placeholders in the .pptx Template

Put these strings inside text boxes on your PowerPoint word slide:

```
{{Word}}          {{Plural}}        {{Stress}}
{{Vowels}}        {{Hint}}          {{Trap}}
{{Type}}          {{Rule}}          {{Usage}}
{{English}}       {{Pronunciation}}
{{Image}}         {{Audio}}
{{Index}}   ← on the Index slide only
```

- Each placeholder should be the **sole text** in its text box.
- `{{Image}}` — replaced by the word's image at the same position/size.
  Shape is silently removed if no image is provided.
- `{{Audio}}` — replaced by a clickable speaker icon.
  Shape is silently removed if no audio is provided.

---

## Index Line Format Tokens

Configure per template. Available tokens:

| Token      | Value |
|------------|-------|
| `{n}`      | 1-based entry number |
| `{word}`   | the word |
| `{plural}` | plural form |
| `{type}`   | word type |
| `{english}`| English translation |

Example: `{n}. {word} / {plural} ({type})`

---

## Excel Word List Format

Row 1 = headers (skipped). Data starts at row 2.

| Col | Field         | Required |
|-----|---------------|----------|
| 1   | Word          | Yes |
| 2   | Plural        | |
| 3   | Stress        | |
| 4   | Vowels        | |
| 5   | Hint          | |
| 6   | Trap          | |
| 7   | Type          | |
| 8   | Rule          | |
| 9   | Usage         | |
| 10  | English       | |
| 11  | Pronunciation | |
| 12  | Image         | |
| 13  | Audio         | |

### Asset Columns (12 & 13) — three options:

1. **Empty** — no asset; `{{Image}}` / `{{Audio}}` shape removed from slide.
2. **File path** — absolute or relative to the Excel file location.
3. **Embedded** — insert via *Insert → Pictures → Place in Cell* (images)
   or *Insert → Object* (audio). Leave cell text empty.

You can also click any Image/Audio cell in the app preview to browse for a file.

---

## Saved Configuration Location

```
%APPDATA%\SlideshowGenerator\templates\
```

---

## NuGet Packages

| Package | Version | Purpose |
|---------|---------|---------|
| ClosedXML | 0.102.2 | Read Excel files + embedded pictures |
| DocumentFormat.OpenXml | 3.1.0 | Read/write PowerPoint XML |
