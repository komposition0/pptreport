using System.Text.Json;
using System.Text.Json.Serialization;
using ReportGenerator.Core;

namespace ReportGenerator.Storage;

/// <summary>Reads and writes report projects as JSON.</summary>
public sealed class ReportProjectStore
{
    private static readonly JsonSerializerOptions SerializerOptions = new()
    {
        WriteIndented = true,
        PropertyNameCaseInsensitive = true,
        Converters = { new JsonStringEnumConverter() }
    };

    /// <summary>Saves a report project to the specified JSON file.</summary>
    public async Task SaveAsync(string path, ReportProject project, CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        ArgumentNullException.ThrowIfNull(project);

        project.ModifiedUtc = DateTime.UtcNow;
        var destination = Path.GetFullPath(path);
        Directory.CreateDirectory(Path.GetDirectoryName(destination)!);
        await using var stream = File.Create(destination);
        await JsonSerializer.SerializeAsync(stream, project, SerializerOptions, cancellationToken);
    }

    /// <summary>Loads a report project from a JSON file.</summary>
    public async Task<ReportProject> LoadAsync(string path, CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        await using var stream = File.OpenRead(path);
        return await JsonSerializer.DeserializeAsync<ReportProject>(stream, SerializerOptions, cancellationToken)
            ?? throw new InvalidDataException("The report project file is empty or invalid.");
    }
}
