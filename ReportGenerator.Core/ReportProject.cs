namespace ReportGenerator.Core;

/// <summary>Contains the saved data required to resume or generate a report.</summary>
public sealed class ReportProject
{
    /// <summary>Gets or sets the report's display name.</summary>
    public string ReportName { get; set; } = "Untitled Report";

    /// <summary>Gets or sets the absolute path of the source template.</summary>
    public string TemplatePath { get; set; } = string.Empty;

    /// <summary>Gets or sets the SHA-256 fingerprint of the source template.</summary>
    public string TemplateFingerprint { get; set; } = string.Empty;

    /// <summary>Gets or sets the report values keyed by placeholder name.</summary>
    public Dictionary<string, string> Values { get; set; } = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Gets or sets image scaling modes keyed by image placeholder name.</summary>
    public Dictionary<string, ImageScalingMode> ImageScaling { get; set; } = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Gets or sets the creation time in UTC.</summary>
    public DateTime CreatedUtc { get; set; } = DateTime.UtcNow;

    /// <summary>Gets or sets the last modification time in UTC.</summary>
    public DateTime ModifiedUtc { get; set; } = DateTime.UtcNow;
}
