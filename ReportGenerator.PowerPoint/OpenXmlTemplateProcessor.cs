using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using ReportGenerator.Core;

namespace ReportGenerator.PowerPoint;

/// <summary>Processes PowerPoint templates through the Open XML SDK.</summary>
public sealed class OpenXmlTemplateProcessor : ITemplateProcessor
{
    private static readonly Regex PlaceholderPattern = new(@"\{\{(?<name>[A-Za-z0-9_.-]+)\}\}", RegexOptions.Compiled);

    /// <inheritdoc />
    public IReadOnlyList<PlaceholderDefinition> DiscoverPlaceholders(string templatePath)
    {
        ValidateTemplatePath(templatePath);

        using var document = PresentationDocument.Open(templatePath, false);
        var definitions = new Dictionary<string, PlaceholderDefinition>(StringComparer.OrdinalIgnoreCase);
        var slideNumber = 0;

        foreach (var slidePart in GetSlideParts(document))
        {
            slideNumber++;
            foreach (var paragraph in slidePart.Slide.Descendants<A.Paragraph>())
            {
                var text = GetParagraphText(paragraph);
                var shapeId = paragraph.Ancestors<P.Shape>().FirstOrDefault()
                    ?.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value
                    ?? paragraph.Ancestors<P.GraphicFrame>().FirstOrDefault()
                        ?.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Id?.Value
                    ?? 0;
                foreach (Match match in PlaceholderPattern.Matches(text))
                {
                    var name = match.Groups["name"].Value;
                    if (definitions.ContainsKey(name))
                    {
                        continue;
                    }

                    var isImage = name.StartsWith("IMAGE_", StringComparison.OrdinalIgnoreCase);
                    definitions.Add(name, new PlaceholderDefinition
                    {
                        Name = name,
                        DisplayName = ToDisplayName(name),
                        Type = isImage ? PlaceholderType.Image : PlaceholderType.Text,
                        SlideNumber = slideNumber,
                        ShapeId = shapeId,
                        Required = false
                    });
                }
            }
        }

        return definitions.Values.OrderBy(item => item.SlideNumber).ThenBy(item => item.Name).ToArray();
    }

    /// <inheritdoc />
    public IReadOnlyList<SlideSummary> GetSlideSummaries(string templatePath)
    {
        ValidateTemplatePath(templatePath);

        using var document = PresentationDocument.Open(templatePath, false);
        var slides = new List<SlideSummary>();
        var number = 0;
        foreach (var slidePart in GetSlideParts(document))
        {
            number++;
            var slideText = slidePart.Slide.Descendants<A.Paragraph>()
                .Select(GetParagraphText)
                .Where(text => !string.IsNullOrWhiteSpace(text));
            slides.Add(new SlideSummary { Number = number, Text = string.Join(Environment.NewLine, slideText) });
        }

        return slides;
    }

    /// <inheritdoc />
    public Task GenerateAsync(
        string templatePath,
        string outputPath,
        IReadOnlyDictionary<string, string> values,
        IReadOnlyDictionary<string, ImageScalingMode> imageScaling,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(values);
        ArgumentNullException.ThrowIfNull(imageScaling);
        ValidateTemplatePath(templatePath);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);

        var source = Path.GetFullPath(templatePath);
        var destination = Path.GetFullPath(outputPath);
        if (string.Equals(source, destination, StringComparison.OrdinalIgnoreCase))
        {
            throw new ArgumentException("The generated presentation must not overwrite its source template.", nameof(outputPath));
        }

        return Task.Run(() =>
        {
            cancellationToken.ThrowIfCancellationRequested();
            Directory.CreateDirectory(Path.GetDirectoryName(destination)!);
            File.Copy(source, destination, true);

            using var document = PresentationDocument.Open(destination, true);
            foreach (var slidePart in GetSlideParts(document))
            {
                cancellationToken.ThrowIfCancellationRequested();
                var shapes = slidePart.Slide.Descendants<P.Shape>().ToArray();
                foreach (var shape in shapes)
                {
                    var shapeText = GetShapeText(shape);
                    var imageMatch = PlaceholderPattern.Matches(shapeText)
                        .Cast<Match>()
                        .FirstOrDefault(match => match.Groups["name"].Value.StartsWith("IMAGE_", StringComparison.OrdinalIgnoreCase));

                    if (imageMatch is not null)
                    {
                        var imageName = imageMatch.Groups["name"].Value;
                        if (values.TryGetValue(imageName, out var imagePath) && !string.IsNullOrWhiteSpace(imagePath))
                        {
                            ReplaceImage(slidePart, shape, imagePath, GetScalingMode(imageScaling, imageName));
                            continue;
                        }
                    }

                }

                foreach (var paragraph in slidePart.Slide.Descendants<A.Paragraph>().ToArray())
                {
                    ReplaceTextInParagraph(paragraph, values);
                }
            }

            document.PresentationPart!.Presentation.Save();
        }, cancellationToken);
    }

    /// <summary>Computes the SHA-256 fingerprint used to identify a template version.</summary>
    public static string GetTemplateFingerprint(string templatePath)
    {
        ValidateTemplatePath(templatePath);
        using var stream = File.OpenRead(templatePath);
        return Convert.ToHexString(SHA256.HashData(stream));
    }

    private static IEnumerable<SlidePart> GetSlideParts(PresentationDocument document)
    {
        var presentationPart = document.PresentationPart
            ?? throw new InvalidDataException("The PowerPoint file has no presentation part.");
        var slideIds = presentationPart.Presentation.SlideIdList
            ?? throw new InvalidDataException("The PowerPoint file contains no slides.");
        foreach (var slideId in slideIds.Elements<P.SlideId>())
        {
            if (presentationPart.GetPartById(slideId.RelationshipId!) is not SlidePart slidePart)
            {
                throw new InvalidDataException("The presentation contains an invalid slide relationship.");
            }

            yield return slidePart;
        }
    }

    private static void ValidateTemplatePath(string templatePath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(templatePath);
        if (!File.Exists(templatePath))
        {
            throw new FileNotFoundException("PowerPoint template was not found.", templatePath);
        }

        if (!string.Equals(Path.GetExtension(templatePath), ".pptx", StringComparison.OrdinalIgnoreCase))
        {
            throw new ArgumentException("Only .pptx templates are supported.", nameof(templatePath));
        }
    }

    private static string GetShapeText(P.Shape shape) =>
        string.Join(Environment.NewLine, shape.TextBody?.Descendants<A.Paragraph>().Select(GetParagraphText) ?? Enumerable.Empty<string>());

    private static string GetParagraphText(A.Paragraph paragraph) =>
        string.Concat(paragraph.Descendants<A.Text>().Select(text => text.Text));

    private static string ToDisplayName(string name) =>
        string.Join(" ", name.Split('_', StringSplitOptions.RemoveEmptyEntries)
            .Select(part => part.Length == 0 ? part : char.ToUpperInvariant(part[0]) + part[1..].ToLowerInvariant()));

    private static ImageScalingMode GetScalingMode(
        IReadOnlyDictionary<string, ImageScalingMode> imageScaling,
        string name) =>
        imageScaling.TryGetValue(name, out var mode) ? mode : ImageScalingMode.Fit;

    private static void ReplaceTextInParagraph(A.Paragraph paragraph, IReadOnlyDictionary<string, string> values)
    {
        var textNodes = paragraph.Descendants<A.Text>().ToArray();
        if (textNodes.Length == 0)
        {
            return;
        }

        var originalValues = textNodes.Select(node => node.Text ?? string.Empty).ToArray();
        var combined = string.Concat(originalValues);
        var matches = PlaceholderPattern.Matches(combined).Cast<Match>().ToArray();
        if (matches.Length == 0)
        {
            return;
        }

        var nodeStarts = new int[textNodes.Length];
        for (var i = 1; i < textNodes.Length; i++)
        {
            nodeStarts[i] = nodeStarts[i - 1] + originalValues[i - 1].Length;
        }

        for (var nodeIndex = 0; nodeIndex < textNodes.Length; nodeIndex++)
        {
            var nodeStart = nodeStarts[nodeIndex];
            var nodeEnd = nodeStart + originalValues[nodeIndex].Length;
            var output = new StringBuilder();
            var cursor = nodeStart;

            foreach (var match in matches)
            {
                var matchStart = match.Index;
                var matchEnd = match.Index + match.Length;
                var overlapStart = Math.Max(nodeStart, matchStart);
                var overlapEnd = Math.Min(nodeEnd, matchEnd);
                if (overlapStart >= overlapEnd)
                {
                    continue;
                }

                if (overlapStart > cursor)
                {
                    output.Append(combined, cursor, overlapStart - cursor);
                }

                if (overlapStart == matchStart && values.TryGetValue(match.Groups["name"].Value, out var replacement))
                {
                    output.Append(replacement);
                }

                cursor = overlapEnd;
            }

            if (cursor < nodeEnd)
            {
                output.Append(combined, cursor, nodeEnd - cursor);
            }

            textNodes[nodeIndex].Text = output.ToString();
            if (output.Length > 0 && (char.IsWhiteSpace(output[0]) || char.IsWhiteSpace(output[^1])))
            {
                textNodes[nodeIndex].SetAttribute(new OpenXmlAttribute(
                    "xml",
                    "space",
                    "http://www.w3.org/XML/1998/namespace",
                    "preserve"));
            }
        }
    }

    private static void ReplaceImage(SlidePart slidePart, P.Shape shape, string imagePath, ImageScalingMode scalingMode)
    {
        if (!File.Exists(imagePath))
        {
            throw new FileNotFoundException("Image placeholder file was not found.", imagePath);
        }

        var transform = shape.ShapeProperties?.GetFirstChild<A.Transform2D>()
            ?? throw new InvalidDataException($"Image placeholder shape {shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value} has no position.");
        var offset = transform.Offset ?? throw new InvalidDataException("Image placeholder has no position.");
        var extents = transform.Extents ?? throw new InvalidDataException("Image placeholder has no dimensions.");
        var width = extents.Cx!.Value;
        var height = extents.Cy!.Value;
        var (pixelWidth, pixelHeight) = ReadImageDimensions(imagePath);

        var targetWidth = width;
        var targetHeight = height;
        var x = offset.X!.Value;
        var y = offset.Y!.Value;
        var sourceRectangle = new A.SourceRectangle();

        if (scalingMode == ImageScalingMode.Fit || scalingMode == ImageScalingMode.OriginalSize)
        {
            var desiredWidth = scalingMode == ImageScalingMode.OriginalSize ? pixelWidth * 9525d : width;
            var desiredHeight = scalingMode == ImageScalingMode.OriginalSize ? pixelHeight * 9525d : height;
            var scale = Math.Min(1d, Math.Min(width / desiredWidth, height / desiredHeight));
            if (scalingMode == ImageScalingMode.Fit)
            {
                scale = Math.Min(width / (double)pixelWidth, height / (double)pixelHeight);
            }

            var unitScale = scalingMode == ImageScalingMode.OriginalSize ? 9525d : 1d;
            targetWidth = (long)Math.Round(pixelWidth * unitScale * scale);
            targetHeight = (long)Math.Round(pixelHeight * unitScale * scale);
            x += (width - targetWidth) / 2;
            y += (height - targetHeight) / 2;
        }
        else if (scalingMode == ImageScalingMode.Fill)
        {
            var imageAspect = pixelWidth / (double)pixelHeight;
            var boxAspect = width / (double)height;
            if (imageAspect > boxAspect)
            {
                var crop = (int)Math.Round((1d - boxAspect / imageAspect) * 50000);
                sourceRectangle.Left = crop;
                sourceRectangle.Right = crop;
            }
            else if (imageAspect < boxAspect)
            {
                var crop = (int)Math.Round((1d - imageAspect / boxAspect) * 50000);
                sourceRectangle.Top = crop;
                sourceRectangle.Bottom = crop;
            }
        }

        var imagePart = slidePart.AddImagePart(GetImagePartType(imagePath));
        using (var stream = File.OpenRead(imagePath))
        {
            imagePart.FeedData(stream);
        }

        var relationshipId = slidePart.GetIdOfPart(imagePart);
        var shapeId = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value ?? 1U;
        var picture = new P.Picture(
            new P.NonVisualPictureProperties(
                new P.NonVisualDrawingProperties { Id = shapeId, Name = Path.GetFileName(imagePath) },
                new P.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true }),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.BlipFill(
                new A.Blip { Embed = relationshipId },
                sourceRectangle,
                new A.Stretch(new A.FillRectangle())),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = x, Y = y },
                    new A.Extents { Cx = targetWidth, Cy = targetHeight }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));

        shape.Parent!.InsertAfter(picture, shape);
        shape.Remove();
    }

    private static (int Width, int Height) ReadImageDimensions(string path)
    {
        using var stream = File.OpenRead(path);
        Span<byte> header = stackalloc byte[32];
        var read = stream.Read(header);
        var extension = Path.GetExtension(path).ToLowerInvariant();

        if (extension == ".png" && read >= 24)
        {
            return (ReadBigEndianInt32(header[16..20]), ReadBigEndianInt32(header[20..24]));
        }

        if (extension == ".gif" && read >= 10)
        {
            return (header[6] | header[7] << 8, header[8] | header[9] << 8);
        }

        if (extension == ".bmp" && read >= 26)
        {
            return (Math.Abs(BitConverter.ToInt32(header[18..22])), Math.Abs(BitConverter.ToInt32(header[22..26])));
        }

        if (extension is ".jpg" or ".jpeg")
        {
            stream.Position = 2;
            while (stream.Position + 4 <= stream.Length)
            {
                var markerPrefix = stream.ReadByte();
                if (markerPrefix != 0xFF)
                {
                    continue;
                }

                var marker = stream.ReadByte();
                while (marker == 0xFF)
                {
                    marker = stream.ReadByte();
                }

                var lengthHigh = stream.ReadByte();
                var lengthLow = stream.ReadByte();
                if (lengthHigh < 0 || lengthLow < 0)
                {
                    break;
                }

                var segmentLength = lengthHigh << 8 | lengthLow;
                if (marker is >= 0xC0 and <= 0xC3 or >= 0xC5 and <= 0xC7 or >= 0xC9 and <= 0xCB or >= 0xCD and <= 0xCF)
                {
                    if (segmentLength < 7)
                    {
                        break;
                    }

                    var segment = new byte[5];
                    if (stream.Read(segment, 0, segment.Length) != segment.Length)
                    {
                        break;
                    }

                    return (segment[3] << 8 | segment[4], segment[1] << 8 | segment[2]);
                }

                stream.Position += segmentLength - 2;
            }
        }

        throw new InvalidDataException($"Could not read image dimensions for '{Path.GetFileName(path)}'. Use a PNG, JPEG, GIF, or BMP image.");
    }

    private static int ReadBigEndianInt32(ReadOnlySpan<byte> bytes) =>
        bytes[0] << 24 | bytes[1] << 16 | bytes[2] << 8 | bytes[3];

    private static ImagePartType GetImagePartType(string path) =>
        Path.GetExtension(path).ToLowerInvariant() switch
        {
            ".png" => ImagePartType.Png,
            ".jpg" or ".jpeg" => ImagePartType.Jpeg,
            ".gif" => ImagePartType.Gif,
            ".bmp" => ImagePartType.Bmp,
            _ => throw new InvalidDataException("Only PNG, JPEG, GIF, and BMP image files are supported.")
        };
}
