using LanguageCourseSlides.Models;
using System.Text.Json;

namespace LanguageCourseSlides.Services;

public static class ConfigManager
{
    private static readonly string ConfigDir =
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "LanguageCourseSlides",
            "templates");

    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        WriteIndented  = true,
        Converters     = { new System.Text.Json.Serialization.JsonStringEnumConverter() }
    };

    static ConfigManager() => Directory.CreateDirectory(ConfigDir);

    // -----------------------------------------------------------------------
    // CRUD
    // -----------------------------------------------------------------------

    public static List<TemplateConfig> LoadAll()
    {
        return Directory.GetFiles(ConfigDir, "*.json")
            .Select(f =>
            {
                try { return JsonSerializer.Deserialize<TemplateConfig>(File.ReadAllText(f), JsonOpts); }
                catch { return null; }
            })
            .Where(c => c != null)
            .Cast<TemplateConfig>()
            .OrderBy(c => c.ConfigName)
            .ToList();
    }

    public static void Save(TemplateConfig config)
    {
        var path = GetFilePath(config.ConfigName);
        File.WriteAllText(path, JsonSerializer.Serialize(config, JsonOpts));
    }

    public static void Delete(TemplateConfig config)
    {
        var path = GetFilePath(config.ConfigName);
        if (File.Exists(path)) File.Delete(path);
    }

    public static void Rename(TemplateConfig config, string oldName)
    {
        var oldPath = GetFilePath(oldName);
        if (File.Exists(oldPath)) File.Delete(oldPath);
        Save(config);
    }

    // -----------------------------------------------------------------------
    // Deep clone via JSON round-trip
    // -----------------------------------------------------------------------

    public static TemplateConfig Clone(TemplateConfig source) =>
        JsonSerializer.Deserialize<TemplateConfig>(
            JsonSerializer.Serialize(source, JsonOpts), JsonOpts)!;

    // -----------------------------------------------------------------------
    // Helpers
    // -----------------------------------------------------------------------

    private static string GetFilePath(string name)
    {
        var safe = string.Concat(name.Split(Path.GetInvalidFileNameChars()));
        return Path.Combine(ConfigDir, $"{safe}.json");
    }

    public static string ConfigDirectory => ConfigDir;
}
