namespace SlideshowGenerator.Models;

public enum SlideRole
{
    Static,
    Index,
    Word
}

public class SlideDefinition
{
    public int       SlideIndex       { get; set; }
    public SlideRole Role             { get; set; }
    public string    IndexPlaceholder { get; set; } = "{{Index}}";
    public string    Label            { get; set; } = "";
}
