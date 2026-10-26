namespace ReportGenerator.Core;

/// <summary>Defines the supported types of report fields.</summary>
public enum PlaceholderType
{
    /// <summary>A single-line text value.</summary>
    Text,
    /// <summary>A multi-line text value.</summary>
    MultilineText,
    /// <summary>A numeric value.</summary>
    Number,
    /// <summary>A calendar date.</summary>
    Date,
    /// <summary>An image file.</summary>
    Image,
    /// <summary>A true or false value.</summary>
    Boolean,
    /// <summary>A value selected from configured choices.</summary>
    Dropdown
}

/// <summary>Defines an editable placeholder discovered in a presentation template.</summary>
public sealed class PlaceholderDefinition
{
    /// <summary>Gets or sets the unique identifier for this definition.</summary>
    public string Id { get; set; } = Guid.NewGuid().ToString("N");

    /// <summary>Gets or sets the token name without braces.</summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>Gets or sets the label shown to users.</summary>
    public string DisplayName { get; set; } = string.Empty;

    /// <summary>Gets or sets the editor type.</summary>
    public PlaceholderType Type { get; set; } = PlaceholderType.Text;

    /// <summary>Gets or sets the one-based slide number of the first occurrence.</summary>
    public int SlideNumber { get; set; }

    /// <summary>Gets or sets the PowerPoint shape identifier of the first occurrence.</summary>
    public uint ShapeId { get; set; }

    /// <summary>Gets or sets the initial value for the field.</summary>
    public string DefaultValue { get; set; } = string.Empty;

    /// <summary>Gets or sets whether a value is required before generation.</summary>
    public bool Required { get; set; }

    /// <summary>Gets or sets an optional description for the field.</summary>
    public string Description { get; set; } = string.Empty;
}

/// <summary>Describes a simplified, text-based slide preview.</summary>
public sealed class SlideSummary
{
    /// <summary>Gets or sets the one-based slide number.</summary>
    public int Number { get; set; }

    /// <summary>Gets or sets the extracted text for the slide.</summary>
    public string Text { get; set; } = string.Empty;
}

/// <summary>Controls how an image is fitted to a PowerPoint image placeholder.</summary>
public enum ImageScalingMode
{
    /// <summary>Preserve the aspect ratio and fit the complete image inside the placeholder.</summary>
    Fit,
    /// <summary>Preserve the aspect ratio and crop excess image content.</summary>
    Fill,
    /// <summary>Fill the entire placeholder, allowing distortion.</summary>
    Stretch,
    /// <summary>Use the image's native pixel dimensions, constrained to the placeholder.</summary>
    OriginalSize
}
