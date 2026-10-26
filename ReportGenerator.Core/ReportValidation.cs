using System.Globalization;

namespace ReportGenerator.Core;

/// <summary>Represents a validation issue that prevents report generation.</summary>
public sealed record ValidationIssue(string PlaceholderName, string Message);

/// <summary>Validates required report values and image references.</summary>
public static class ReportValidation
{
    /// <summary>Returns all validation issues found in the supplied report values.</summary>
    public static IReadOnlyList<ValidationIssue> Validate(
        IEnumerable<PlaceholderDefinition> definitions,
        IReadOnlyDictionary<string, string> values)
    {
        ArgumentNullException.ThrowIfNull(definitions);
        ArgumentNullException.ThrowIfNull(values);

        var issues = new List<ValidationIssue>();
        foreach (var definition in definitions)
        {
            values.TryGetValue(definition.Name, out var value);
            value ??= definition.DefaultValue;

            if (definition.Required && string.IsNullOrWhiteSpace(value))
            {
                issues.Add(new ValidationIssue(definition.Name, $"{definition.DisplayName} is required."));
                continue;
            }

            if (string.IsNullOrWhiteSpace(value))
            {
                continue;
            }

            if (definition.Type == PlaceholderType.Number &&
                !decimal.TryParse(value, NumberStyles.Number, CultureInfo.CurrentCulture, out _))
            {
                issues.Add(new ValidationIssue(definition.Name, $"{definition.DisplayName} must be a valid number."));
            }
            else if (definition.Type == PlaceholderType.Date &&
                     !DateTime.TryParse(value, CultureInfo.CurrentCulture, DateTimeStyles.None, out _))
            {
                issues.Add(new ValidationIssue(definition.Name, $"{definition.DisplayName} must be a valid date."));
            }
            else if (definition.Type == PlaceholderType.Image && !File.Exists(value))
            {
                issues.Add(new ValidationIssue(definition.Name, $"Image file not found: {value}"));
            }
            else if (definition.Type == PlaceholderType.Image &&
                     !new[] { ".png", ".jpg", ".jpeg", ".gif", ".bmp" }
                         .Contains(Path.GetExtension(value), StringComparer.OrdinalIgnoreCase))
            {
                issues.Add(new ValidationIssue(definition.Name, $"{definition.DisplayName} must be a PNG, JPEG, GIF, or BMP image."));
            }
        }

        return issues;
    }
}
