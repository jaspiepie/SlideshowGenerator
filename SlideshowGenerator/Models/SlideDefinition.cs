namespace LanguageCourseSlides.Models;

public enum SlideRole { Static, Index, Word }

public class SlideDefinition
{
    public int       SlideIndex { get; set; }
    public SlideRole Role       { get; set; }
    public string    Label      { get; set; } = "";
}
