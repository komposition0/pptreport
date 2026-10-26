# PowerPoint Report Generator

A Windows Forms application for filling existing `.pptx` templates without requiring Microsoft PowerPoint. The solution targets .NET 8 and uses the Open XML SDK to preserve the original presentation structure.

## Build and run

Open `PowerPointReportGenerator.sln` in Visual Studio 2022, or use the .NET SDK:

```powershell
dotnet build PowerPointReportGenerator.sln
dotnet run --project ReportGenerator.App
```

Load a `.pptx` template, edit discovered placeholders in the right-hand panel, optionally select images and scaling modes, save the report as JSON, and generate a new PowerPoint file. The source template is never overwritten. The center panel is a clearly labeled simplified text preview; it does not render PowerPoint slide visuals.

## Example template

Create the included five-slide engineering report template:

```powershell
dotnet run --project ReportGenerator.App -- --create-sample-template Samples\StandardReportTemplate.pptx
```

Then load `Samples\StandardReportTemplate.pptx` from the application. The sample includes report information, test results, image locations, and conclusion placeholders.

Placeholders use `{{NAME}}` syntax. Names beginning with `IMAGE_` are presented as image fields. Save report projects as JSON to restore their template reference, entered values, and image scaling preferences.

## Projects

- `ReportGenerator.App` — WinForms user interface.
- `ReportGenerator.Core` — report data models and validation.
- `ReportGenerator.PowerPoint` — Open XML template discovery and generation.
- `ReportGenerator.Storage` — JSON report project persistence.
- `ReportGenerator.Tests` — focused unit and integration tests.

## Current limitations

- Slide previews are text-only; actual PowerPoint slide rendering is not performed.
- Template-specific field configuration, project packaging, PDF export, and advanced batch features are not part of this MVP.
- Image placeholders are shapes containing `{{IMAGE_NAME}}`; supported input formats are PNG, JPEG, GIF, and BMP.
