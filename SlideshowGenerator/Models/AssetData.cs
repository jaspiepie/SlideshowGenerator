namespace LanguageCourseSlides.Models;

/// <summary>
/// Carries image or audio data from either a file path or bytes
/// embedded directly inside the Excel workbook.
/// </summary>
public class AssetData
{
    public string?  FilePath    { get; init; }
    public byte[]?  Bytes       { get; init; }
    public string   Extension   { get; init; } = "";
    public string   ContentType { get; init; } = "";

    public bool IsValid =>
        (FilePath != null && File.Exists(FilePath)) ||
        (Bytes    != null && Bytes.Length > 0);

    /// <summary>Opens a stream over the asset regardless of its origin.</summary>
    public Stream OpenStream() =>
        FilePath != null
            ? File.OpenRead(FilePath)
            : new MemoryStream(Bytes!);

    // ------------------------------------------------------------------
    // Factory helpers
    // ------------------------------------------------------------------

    public static AssetData? FromPath(string? raw, string baseDir)
    {
        if (string.IsNullOrWhiteSpace(raw)) return null;

        var path = Path.IsPathRooted(raw)
            ? raw
            : Path.GetFullPath(Path.Combine(baseDir, raw));

        if (!File.Exists(path)) return null;

        var ext = Path.GetExtension(path).ToLowerInvariant();
        return new AssetData
        {
            FilePath    = path,
            Extension   = ext,
            ContentType = ResolveContentType(ext),
        };
    }

    public static AssetData? FromBytes(byte[]? bytes, string extension)
    {
        if (bytes == null || bytes.Length == 0) return null;
        var ext = extension.ToLowerInvariant();
        return new AssetData
        {
            Bytes       = bytes,
            Extension   = ext,
            ContentType = ResolveContentType(ext),
        };
    }

    public static string ResolveContentType(string ext) => ext switch
    {
        ".jpg" or ".jpeg" => "image/jpeg",
        ".png"            => "image/png",
        ".gif"            => "image/gif",
        ".webp"           => "image/webp",
        ".bmp"            => "image/bmp",
        ".mp3"            => "audio/mpeg",
        ".wav"            => "audio/wav",
        ".m4a"            => "audio/mp4",
        ".ogg"            => "audio/ogg",
        _                 => "application/octet-stream"
    };
}
