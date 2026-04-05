using System.IO;
using System.Text.Json.Serialization;

namespace SlideshowGenerator.Models;

public class WordEntry
{
    public string Word          { get; set; } = "";
    public string Plural        { get; set; } = "";
    public string Stress        { get; set; } = "";
    public string Vowels        { get; set; } = "";
    public string Hint          { get; set; } = "";
    public string Trap          { get; set; } = "";
    public string Type          { get; set; } = "";
    public string Rule          { get; set; } = "";
    public string Usage         { get; set; } = "";
    public string English       { get; set; } = "";
    public string Pronunciation { get; set; } = "";

    // Populated by ExcelReader — null means no asset was found
    [JsonIgnore] public AssetData? Image { get; set; }
    [JsonIgnore] public AssetData? Audio { get; set; }

    // Display-only strings for the preview grid
    public string ImageStatus => Image?.IsValid == true
        ? (Image.FilePath != null ? $"📁 {Path.GetFileName(Image.FilePath)}" : "📎 Embedded")
        : "";

    public string AudioStatus => Audio?.IsValid == true
        ? (Audio.FilePath != null ? $"📁 {Path.GetFileName(Audio.FilePath)}" : "📎 Embedded")
        : "";

    public bool HasImage => Image?.IsValid == true;
    public bool HasAudio => Audio?.IsValid == true;

    public Dictionary<string, string> ToPlaceholders() => new()
    {
        ["{{Word}}"]          = Word,
        ["{{Plural}}"]        = Plural,
        ["{{Stress}}"]        = Stress,
        ["{{Vowels}}"]        = Vowels,
        ["{{Hint}}"]          = Hint,
        ["{{Trap}}"]          = Trap,
        ["{{Type}}"]          = Type,
        ["{{Rule}}"]          = Rule,
        ["{{Usage}}"]         = Usage,
        ["{{English}}"]       = English,
        ["{{Pronunciation}}"] = Pronunciation,
    };
}
