using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ReportGenerator.PowerPoint;

/// <summary>Creates the five-slide example template used to explore the application.</summary>
public static class SampleTemplateFactory
{
    /// <summary>Creates a sample report template at the requested .pptx path.</summary>
    public static void Create(string path)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        if (!string.Equals(System.IO.Path.GetExtension(path), ".pptx", StringComparison.OrdinalIgnoreCase))
        {
            throw new ArgumentException("The sample template path must end in .pptx.", nameof(path));
        }

        var destination = System.IO.Path.GetFullPath(path);
        Directory.CreateDirectory(System.IO.Path.GetDirectoryName(destination)!);
        using var document = PresentationDocument.Create(destination, PresentationDocumentType.Presentation);
        var presentationPart = document.AddPresentationPart();
        var masterPart = presentationPart.AddNewPart<SlideMasterPart>();
        var layoutPart = masterPart.AddNewPart<SlideLayoutPart>();
        layoutPart.SlideLayout = new P.SlideLayout(
            new P.CommonSlideData(CreateShapeTree(Array.Empty<(string, string, long, long, long, long)>())),
            new P.ColorMapOverride(new A.MasterColorMapping()))
        {
            Type = P.SlideLayoutValues.Blank
        };
        layoutPart.SlideLayout.Save();
        masterPart.SlideMaster = new P.SlideMaster(
            new P.CommonSlideData(CreateShapeTree(Array.Empty<(string, string, long, long, long, long)>())),
            new P.ColorMap
            {
                Background1 = A.ColorSchemeIndexValues.Light1,
                Text1 = A.ColorSchemeIndexValues.Dark1,
                Background2 = A.ColorSchemeIndexValues.Light2,
                Text2 = A.ColorSchemeIndexValues.Dark2,
                Accent1 = A.ColorSchemeIndexValues.Accent1,
                Accent2 = A.ColorSchemeIndexValues.Accent2,
                Accent3 = A.ColorSchemeIndexValues.Accent3,
                Accent4 = A.ColorSchemeIndexValues.Accent4,
                Accent5 = A.ColorSchemeIndexValues.Accent5,
                Accent6 = A.ColorSchemeIndexValues.Accent6,
                Hyperlink = A.ColorSchemeIndexValues.Hyperlink,
                FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink
            },
            new P.SlideLayoutIdList(new P.SlideLayoutId
            {
                Id = 2147483648U,
                RelationshipId = masterPart.GetIdOfPart(layoutPart)
            }),
            new P.TextStyles(new P.TitleStyle(), new P.BodyStyle(), new P.OtherStyle()));
        masterPart.SlideMaster.Save();

        var slideIds = new P.SlideIdList();
        var slides = new[]
        {
            new[]
            {
                ("Report Title", "{{REPORT_TITLE}}", 600000L, 900000L, 10900000L, 900000L),
                ("Report Date", "Date: {{REPORT_DATE}}", 600000L, 2300000L, 10900000L, 500000L),
                ("Project", "Project: {{PROJECT_NAME}}", 600000L, 3100000L, 10900000L, 500000L)
            },
            new[]
            {
                ("Slide Heading", "General Information", 600000L, 450000L, 10900000L, 600000L),
                ("Engineer", "Engineer: {{ENGINEER_NAME}}", 600000L, 1600000L, 10900000L, 500000L),
                ("Device", "Device ID: {{DEVICE_ID}}", 600000L, 2400000L, 10900000L, 500000L),
                ("Lot", "Lot: {{LOT_NUMBER}}", 600000L, 3200000L, 10900000L, 500000L),
                ("Wafer", "Wafer: {{WAFER_NUMBER}}", 600000L, 4000000L, 10900000L, 500000L)
            },
            new[]
            {
                ("Slide Heading", "Test Results", 600000L, 450000L, 10900000L, 600000L),
                ("Summary", "Summary: {{TEST_SUMMARY}}", 600000L, 1500000L, 10900000L, 1200000L),
                ("Results", "Results: {{TEST_RESULTS}}", 600000L, 3000000L, 10900000L, 1200000L),
                ("Conditions", "Conditions: {{TEST_CONDITIONS}}", 600000L, 4700000L, 10900000L, 800000L)
            },
            new[]
            {
                ("Slide Heading", "Test Images", 600000L, 450000L, 10900000L, 600000L),
                ("Setup Image", "{{IMAGE_TEST_SETUP}}", 600000L, 1400000L, 5200000L, 4300000L),
                ("Device Image", "{{IMAGE_DEVICE_PHOTO}}", 6600000L, 1400000L, 5200000L, 4300000L)
            },
            new[]
            {
                ("Slide Heading", "Conclusion", 600000L, 450000L, 10900000L, 600000L),
                ("Conclusion Text", "{{CONCLUSION}}", 600000L, 1600000L, 10900000L, 1400000L),
                ("Recommendations", "Recommendations: {{RECOMMENDATIONS}}", 600000L, 3500000L, 10900000L, 1200000L)
            }
        };

        for (var index = 0; index < slides.Length; index++)
        {
            var slidePart = presentationPart.AddNewPart<SlidePart>();
            slidePart.AddPart(layoutPart);
            slidePart.Slide = new P.Slide(
                new P.CommonSlideData(CreateShapeTree(slides[index])),
                new P.ColorMapOverride(new A.MasterColorMapping()));
            slidePart.Slide.Save();
            slideIds.Append(new P.SlideId
            {
                Id = (uint)(256 + index),
                RelationshipId = presentationPart.GetIdOfPart(slidePart)
            });
        }

        presentationPart.Presentation = new P.Presentation(
            new P.SlideMasterIdList(new P.SlideMasterId
            {
                Id = 2147483648U,
                RelationshipId = presentationPart.GetIdOfPart(masterPart)
            }),
            slideIds,
            new P.SlideSize { Cx = 12192000, Cy = 6858000, Type = P.SlideSizeValues.Screen16x9 },
            new P.NotesSize { Cx = 6858000, Cy = 9144000 });
        presentationPart.Presentation.Save();
    }

    private static P.ShapeTree CreateShapeTree(
        IEnumerable<(string Name, string Text, long X, long Y, long Width, long Height)> shapes)
    {
        var tree = new P.ShapeTree(
            new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                new P.NonVisualGroupShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.GroupShapeProperties(
                new A.TransformGroup(
                    new A.Offset { X = 0L, Y = 0L },
                    new A.Extents { Cx = 0L, Cy = 0L },
                    new A.ChildOffset { X = 0L, Y = 0L },
                    new A.ChildExtents { Cx = 0L, Cy = 0L })));

        uint id = 2;
        foreach (var shape in shapes)
        {
            tree.Append(new P.Shape(
                new P.NonVisualShapeProperties(
                    new P.NonVisualDrawingProperties { Id = id++, Name = shape.Name },
                    new P.NonVisualShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.ShapeProperties(
                    new A.Transform2D(
                        new A.Offset { X = shape.X, Y = shape.Y },
                        new A.Extents { Cx = shape.Width, Cy = shape.Height }),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }),
                new P.TextBody(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(
                        new A.Run(new A.RunProperties(), new A.Text(shape.Text))))));
        }

        return tree;
    }
}
