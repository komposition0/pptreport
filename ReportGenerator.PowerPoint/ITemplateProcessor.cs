using ReportGenerator.Core;

namespace ReportGenerator.PowerPoint;

/// <summary>Discovers template placeholders and generates populated PowerPoint files.</summary>
public interface ITemplateProcessor
{
    /// <summary>Discovers placeholders in a PowerPoint template.</summary>
    IReadOnlyList<PlaceholderDefinition> DiscoverPlaceholders(string templatePath);

    /// <summary>Extracts slide text for a simplified preview.</summary>
    IReadOnlyList<SlideSummary> GetSlideSummaries(string templatePath);

    /// <summary>Creates a populated copy of a template without modifying the source.</summary>
    Task GenerateAsync(
        string templatePath,
        string outputPath,
        IReadOnlyDictionary<string, string> values,
        IReadOnlyDictionary<string, ImageScalingMode> imageScaling,
        CancellationToken cancellationToken = default);
}
