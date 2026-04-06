namespace LanguageCourseSlides.Models;

public class TemplateConfig
{
    public string             ConfigName        { get; set; } = "";
    public string             TemplatePath      { get; set; } = "";
    public List<SlideDefinition> Slides         { get; set; } = [];
    public bool               HyperlinkIndex    { get; set; } = true;
    public string             IndexLineFormat   { get; set; } = "{word} {type} {english}";
    public int                WordsPerIndexPage { get; set; } = 20;
    public DateTime           CreatedAt         { get; set; } = DateTime.Now;

    // ------------------------------------------------------------------
    // Shape name conventions (must match names given in PowerPoint)
    // ------------------------------------------------------------------
    // On WORD slides:
    //   shape_audio   → plays this word's audio (removed if no audio)
    //   shape_image   → replaced by this word's image (removed if no image)
    //   shape_next    → hyperlinks to the next word slide
    //   shape_index   → hyperlinks back to the first index slide
    //
    // On INDEX slides:
    //   {{Index}}     → text placeholder replaced with the word list
    //
    // On any slide, text placeholders {{Word}} {{Plural}} etc. are replaced.
    // ------------------------------------------------------------------

    public SlideDefinition? IndexSlide =>
        Slides.FirstOrDefault(s => s.Role == SlideRole.Index);
    public SlideDefinition? WordSlide =>
        Slides.FirstOrDefault(s => s.Role == SlideRole.Word);
    public List<SlideDefinition> StaticSlides =>
        Slides.Where(s => s.Role == SlideRole.Static).ToList();

    public bool IsValid =>
        !string.IsNullOrWhiteSpace(ConfigName) &&
        !string.IsNullOrWhiteSpace(TemplatePath) &&
        File.Exists(TemplatePath) &&
        Slides.Any(s => s.Role == SlideRole.Word);

    public override string ToString() => ConfigName;
}
