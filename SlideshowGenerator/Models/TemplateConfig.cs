namespace SlideshowGenerator.Models;

public class TemplateConfig
{
    public string             ConfigName       { get; set; } = "";
    public string             TemplatePath     { get; set; } = "";
    public List<SlideDefinition> Slides        { get; set; } = [];
    public bool               HyperlinkIndex   { get; set; } = true;
    public string             IndexLineFormat  { get; set; } = "{n}. {word}";
    public string             ImagePlaceholder { get; set; } = "{{Image}}";
    public string             AudioPlaceholder { get; set; } = "{{Audio}}";
    public DateTime           CreatedAt        { get; set; } = DateTime.Now;

    // Convenience accessors
    public SlideDefinition? IndexSlide =>
        Slides.FirstOrDefault(s => s.Role == SlideRole.Index);

    public SlideDefinition? WordSlide =>
        Slides.FirstOrDefault(s => s.Role == SlideRole.Word);

    public List<SlideDefinition> StaticSlides =>
        Slides.Where(s => s.Role == SlideRole.Static).ToList();

    public bool IsValid =>
        !string.IsNullOrWhiteSpace(ConfigName)   &&
        !string.IsNullOrWhiteSpace(TemplatePath) &&
        File.Exists(TemplatePath)                &&
        Slides.Any(s => s.Role == SlideRole.Word);
}
