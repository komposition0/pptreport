using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using ReportGenerator.Core;
using ReportGenerator.PowerPoint;
using ReportGenerator.Storage;

namespace ReportGenerator.Tests;

public sealed class ReportGeneratorTests
{
    [Fact]
    public void DiscoverPlaceholders_FindsTextSplitAcrossRunsAndImageFields()
    {
        using var template = TestPresentation.Create();
        var processor = new OpenXmlTemplateProcessor();

        var definitions = processor.DiscoverPlaceholders(template.Path);

        Assert.Equal(2, definitions.Count);
        Assert.Contains(definitions, item => item.Name == "REPORT_TITLE" && item.Type == PlaceholderType.Text);
        Assert.Contains(definitions, item => item.Name == "IMAGE_TEST_SETUP" && item.Type == PlaceholderType.Image);
    }

    [Fact]
    public async Task SampleTemplateFactory_CreatesFiveSlideEngineeringTemplate()
    {
        var directory = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        var path = System.IO.Path.Combine(directory, "sample.pptx");
        try
        {
            SampleTemplateFactory.Create(path);

            var processor = new OpenXmlTemplateProcessor();
            Assert.Equal(5, processor.GetSlideSummaries(path).Count);
            Assert.Contains(processor.DiscoverPlaceholders(path), item => item.Name == "REPORT_TITLE");
            Assert.Contains(processor.DiscoverPlaceholders(path), item => item.Name == "IMAGE_DEVICE_PHOTO");
            using var document = PresentationDocument.Open(path, false);
            var errors = new OpenXmlValidator().Validate(document)
                .Select(error => $"{error.Part?.Uri} {error.Path?.XPath}: {error.Description}")
                .ToArray();
            Assert.True(errors.Length == 0, string.Join(Environment.NewLine, errors));

            var imagePath = System.IO.Path.Combine(directory, "sample.png");
            var outputPath = System.IO.Path.Combine(directory, "completed.pptx");
            await File.WriteAllBytesAsync(imagePath, Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+j1XcAAAAASUVORK5CYII="));
            await processor.GenerateAsync(
                path,
                outputPath,
                new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["REPORT_TITLE"] = "Reliability Review",
                    ["IMAGE_TEST_SETUP"] = imagePath
                },
                new Dictionary<string, ImageScalingMode>());
            using var generated = PresentationDocument.Open(outputPath, false);
            var generatedErrors = new OpenXmlValidator().Validate(generated)
                .Select(error => $"{error.Part?.Uri} {error.Path?.XPath}: {error.Description}")
                .ToArray();
            Assert.True(generatedErrors.Length == 0, string.Join(Environment.NewLine, generatedErrors));
            Assert.Contains("Reliability Review", processor.GetSlideSummaries(outputPath)[0].Text);
        }
        finally
        {
            if (Directory.Exists(directory))
            {
                Directory.Delete(directory, true);
            }
        }
    }

    [Fact]
    public async Task GenerateAsync_ReplacesSplitTextAndPreservesSourceTemplate()
    {
        using var template = TestPresentation.Create();
        var outputPath = System.IO.Path.Combine(template.DirectoryPath, "generated.pptx");
        var processor = new OpenXmlTemplateProcessor();

        await processor.GenerateAsync(
            template.Path,
            outputPath,
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["REPORT_TITLE"] = "Reliability Review"
            },
            new Dictionary<string, ImageScalingMode>());

        Assert.True(File.Exists(template.Path));
        Assert.True(File.Exists(outputPath));
        Assert.Contains("{{REPORT_TITLE}}", processor.GetSlideSummaries(template.Path)[0].Text);
        Assert.Contains("Reliability Review", processor.GetSlideSummaries(outputPath)[0].Text);
        Assert.DoesNotContain("{{REPORT_TITLE}}", processor.GetSlideSummaries(outputPath)[0].Text);
    }

    [Fact]
    public async Task GenerateAsync_ReplacesImagePlaceholderWithPicture()
    {
        using var template = TestPresentation.Create();
        var imagePath = System.IO.Path.Combine(template.DirectoryPath, "pixel.png");
        var outputPath = System.IO.Path.Combine(template.DirectoryPath, "with-image.pptx");
        await File.WriteAllBytesAsync(imagePath, Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+j1XcAAAAASUVORK5CYII="));

        await new OpenXmlTemplateProcessor().GenerateAsync(
            template.Path,
            outputPath,
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["IMAGE_TEST_SETUP"] = imagePath
            },
            new Dictionary<string, ImageScalingMode>());

        using var document = PresentationDocument.Open(outputPath, false);
        var slide = Assert.Single(document.PresentationPart!.SlideParts);
        Assert.Single(slide.Slide.Descendants<P.Picture>());
        Assert.DoesNotContain(slide.Slide.Descendants<P.Shape>(), shape =>
            string.Concat(shape.Descendants<A.Text>().Select(text => text.Text)).Contains("IMAGE_TEST_SETUP", StringComparison.Ordinal));
    }

    [Fact]
    public void Validate_ReportsMissingRequiredAndInvalidNumbers()
    {
        var definitions = new[]
        {
            new PlaceholderDefinition
            {
                Name = "REPORT_TITLE",
                DisplayName = "Report Title",
                Required = true
            },
            new PlaceholderDefinition
            {
                Name = "SAMPLE_COUNT",
                DisplayName = "Sample Count",
                Type = PlaceholderType.Number
            }
        };

        var issues = ReportValidation.Validate(definitions, new Dictionary<string, string>
        {
            ["SAMPLE_COUNT"] = "many"
        });

        Assert.Equal(2, issues.Count);
        Assert.Contains(issues, issue => issue.PlaceholderName == "REPORT_TITLE");
        Assert.Contains(issues, issue => issue.PlaceholderName == "SAMPLE_COUNT");
    }

    [Fact]
    public async Task ReportProjectStore_RoundTripsValuesAndScaling()
    {
        var directory = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            var path = System.IO.Path.Combine(directory, "report.json");
            var project = new ReportProject
            {
                ReportName = "Test Report",
                TemplatePath = "template.pptx",
                Values = new Dictionary<string, string> { ["REPORT_TITLE"] = "Lab results" },
                ImageScaling = new Dictionary<string, ImageScalingMode> { ["IMAGE_SETUP"] = ImageScalingMode.Fill }
            };
            var store = new ReportProjectStore();

            await store.SaveAsync(path, project);
            var loaded = await store.LoadAsync(path);

            Assert.Equal(project.ReportName, loaded.ReportName);
            Assert.Equal("Lab results", loaded.Values["REPORT_TITLE"]);
            Assert.Equal(ImageScalingMode.Fill, loaded.ImageScaling["IMAGE_SETUP"]);
        }
        finally
        {
            Directory.Delete(directory, true);
        }
    }

    private sealed class TestPresentation : IDisposable
    {
        private TestPresentation(string directoryPath, string path)
        {
            DirectoryPath = directoryPath;
            Path = path;
        }

        public string DirectoryPath { get; }
        public string Path { get; }

        public static TestPresentation Create()
        {
            var directory = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            var path = System.IO.Path.Combine(directory, "template.pptx");
            using var document = PresentationDocument.Create(path, PresentationDocumentType.Presentation);
            var presentationPart = document.AddPresentationPart();
            var slidePart = presentationPart.AddNewPart<SlidePart>();
            slidePart.Slide = new P.Slide(
                new P.CommonSlideData(CreateShapeTree()),
                new P.ColorMapOverride(new A.MasterColorMapping()));
            slidePart.Slide.Save();
            presentationPart.Presentation = new P.Presentation(
                new P.SlideIdList(new P.SlideId
                {
                    Id = 256U,
                    RelationshipId = presentationPart.GetIdOfPart(slidePart)
                }));
            presentationPart.Presentation.Save();
            return new TestPresentation(directory, path);
        }

        public void Dispose() => Directory.Delete(DirectoryPath, true);

        private static P.ShapeTree CreateShapeTree() =>
            new(
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.GroupShapeProperties(
                    new A.TransformGroup(
                        new A.Offset { X = 0L, Y = 0L },
                        new A.Extents { Cx = 0L, Cy = 0L },
                        new A.ChildOffset { X = 0L, Y = 0L },
                        new A.ChildExtents { Cx = 0L, Cy = 0L })),
                CreateTextShape(2U, "Title", new[] { "{{REPORT_", "TITLE}}" }),
                CreateTextShape(3U, "Setup", new[] { "{{IMAGE_TEST_SETUP}}" }));

        private static P.Shape CreateTextShape(uint id, string name, IReadOnlyList<string> runs) =>
            new(
                new P.NonVisualShapeProperties(
                    new P.NonVisualDrawingProperties { Id = id, Name = name },
                    new P.NonVisualShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.ShapeProperties(
                    new A.Transform2D(
                        new A.Offset { X = 100000L, Y = 100000L },
                        new A.Extents { Cx = 3000000L, Cy = 500000L }),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }),
                new P.TextBody(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(runs.Select(run =>
                        new A.Run(new A.RunProperties(), new A.Text(run))).ToArray())));
    }
}
