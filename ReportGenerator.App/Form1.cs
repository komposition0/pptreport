using System.Drawing;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using ReportGenerator.Core;
using ReportGenerator.PowerPoint;
using ReportGenerator.Storage;

namespace ReportGenerator.App;

/// <summary>Main window for creating and generating PowerPoint reports.</summary>
public sealed class Form1 : Form
{
    private readonly ITemplateProcessor _templateProcessor = new OpenXmlTemplateProcessor();
    private readonly ReportProjectStore _projectStore = new();
    private readonly List<PlaceholderDefinition> _definitions = [];
    private readonly Dictionary<string, Func<string>> _valueReaders = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, ImageScalingMode> _imageScaling = new(StringComparer.OrdinalIgnoreCase);
    private readonly TabControl _navigationTabs = new();
    private readonly ListBox _slidesList = new();
    private readonly ListBox _fieldsList = new();
    private readonly RichTextBox _preview = new();
    private readonly FlowLayoutPanel _propertyEditor = new();
    private readonly ToolStripStatusLabel _statusLabel = new();
    private ReportProject _project = new();
    private string? _projectFilePath;
    private IReadOnlyList<SlideSummary> _slides = [];
    private int _selectedSlide = 1;
    private bool _buildingControls;

    /// <summary>Initializes the report generator window.</summary>
    public Form1()
    {
        Text = "PowerPoint Report Generator";
        MinimumSize = new Size(1100, 680);
        Size = new Size(1440, 900);
        StartPosition = FormStartPosition.CenterScreen;
        Font = new Font("Segoe UI", 9F);
        BackColor = Color.FromArgb(245, 247, 250);
        BuildLayout();
        StartNewReport();
    }

    private void BuildLayout()
    {
        var toolbar = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 58,
            Padding = new Padding(14, 10, 10, 8),
            BackColor = Color.FromArgb(30, 41, 59),
            WrapContents = false
        };
        AddToolbarButton(toolbar, "New Report", StartNewReport);
        AddToolbarButton(toolbar, "Open Report", OpenReport);
        AddToolbarButton(toolbar, "Load Template", LoadTemplate);
        AddToolbarButton(toolbar, "Save Report", SaveReport);
        AddToolbarButton(toolbar, "Save As", SaveReportAs);
        AddToolbarButton(toolbar, "Generate PowerPoint", GeneratePresentation, true);

        var main = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 3,
            RowCount = 1,
            Padding = new Padding(12),
            BackColor = BackColor
        };
        main.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 23));
        main.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 46));
        main.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 31));

        var leftPanel = CreatePanel();
        leftPanel.Padding = new Padding(10);
        var navigationTitle = CreateSectionTitle("Navigation");
        _navigationTabs.Dock = DockStyle.Fill;
        _slidesList.Dock = DockStyle.Fill;
        _fieldsList.Dock = DockStyle.Fill;
        _slidesList.IntegralHeight = false;
        _fieldsList.IntegralHeight = false;
        _slidesList.SelectedIndexChanged += (_, _) => SelectSlide();
        _fieldsList.SelectedIndexChanged += (_, _) => SelectFieldFromNavigation();
        _navigationTabs.TabPages.Add(new TabPage("Slides"));
        _navigationTabs.TabPages[0].Controls.Add(_slidesList);
        _navigationTabs.TabPages.Add(new TabPage("Placeholders"));
        _navigationTabs.TabPages[1].Controls.Add(_fieldsList);
        leftPanel.Controls.Add(_navigationTabs);
        leftPanel.Controls.Add(navigationTitle);

        var previewPanel = CreatePanel();
        previewPanel.Padding = new Padding(14);
        var previewHeader = CreateSectionTitle("Slide Preview");
        var previewNote = new Label
        {
            Dock = DockStyle.Top,
            Height = 25,
            Text = "Simplified text preview - PowerPoint rendering is not available.",
            ForeColor = Color.FromArgb(100, 116, 139),
            TextAlign = ContentAlignment.MiddleLeft
        };
        _preview.Dock = DockStyle.Fill;
        _preview.ReadOnly = true;
        _preview.BorderStyle = BorderStyle.None;
        _preview.BackColor = Color.White;
        _preview.Font = new Font("Consolas", 11F);
        _preview.DetectUrls = false;
        previewPanel.Controls.Add(_preview);
        previewPanel.Controls.Add(previewNote);
        previewPanel.Controls.Add(previewHeader);

        var propertiesPanel = CreatePanel();
        propertiesPanel.Padding = new Padding(14);
        var propertiesHeader = CreateSectionTitle("Report Properties");
        _propertyEditor.Dock = DockStyle.Fill;
        _propertyEditor.FlowDirection = FlowDirection.TopDown;
        _propertyEditor.WrapContents = false;
        _propertyEditor.AutoScroll = true;
        _propertyEditor.Padding = new Padding(4, 10, 8, 10);
        propertiesPanel.Controls.Add(_propertyEditor);
        propertiesPanel.Controls.Add(propertiesHeader);

        main.Controls.Add(leftPanel, 0, 0);
        main.Controls.Add(previewPanel, 1, 0);
        main.Controls.Add(propertiesPanel, 2, 0);

        var statusStrip = new StatusStrip { BackColor = Color.White };
        _statusLabel.Text = "Ready";
        statusStrip.Items.Add(_statusLabel);
        Controls.Add(main);
        Controls.Add(toolbar);
        Controls.Add(statusStrip);
    }

    private static Panel CreatePanel() => new()
    {
        Dock = DockStyle.Fill,
        BackColor = Color.White,
        Margin = new Padding(6),
        BorderStyle = BorderStyle.FixedSingle
    };

    private static Label CreateSectionTitle(string text) => new()
    {
        Dock = DockStyle.Top,
        Height = 38,
        Text = text,
        Font = new Font("Segoe UI Semibold", 12F, FontStyle.Bold),
        ForeColor = Color.FromArgb(30, 41, 59),
        TextAlign = ContentAlignment.MiddleLeft
    };

    private static void AddToolbarButton(FlowLayoutPanel toolbar, string text, Action action, bool primary = false)
    {
        var button = new Button
        {
            Text = text,
            AutoSize = true,
            Height = 34,
            FlatStyle = FlatStyle.Flat,
            ForeColor = Color.White,
            BackColor = primary ? Color.FromArgb(37, 99, 235) : Color.FromArgb(51, 65, 85),
            Margin = new Padding(0, 0, 8, 0),
            Padding = new Padding(10, 0, 10, 0),
            Cursor = Cursors.Hand
        };
        button.FlatAppearance.BorderSize = 0;
        button.Click += (_, _) => action();
        toolbar.Controls.Add(button);
    }

    private void StartNewReport()
    {
        _project = new ReportProject();
        _projectFilePath = null;
        _definitions.Clear();
        _slides = [];
        _imageScaling.Clear();
        _slidesList.Items.Clear();
        _fieldsList.Items.Clear();
        _propertyEditor.Controls.Clear();
        _preview.Clear();
        _valueReaders.Clear();
        SetStatus("New report");
    }

    private void LoadTemplate()
    {
        using var dialog = new OpenFileDialog
        {
            Title = "Select a PowerPoint template",
            Filter = "PowerPoint presentations (*.pptx)|*.pptx",
            CheckFileExists = true
        };
        if (dialog.ShowDialog(this) != DialogResult.OK)
        {
            return;
        }

        try
        {
            _project.TemplatePath = dialog.FileName;
            _project.TemplateFingerprint = OpenXmlTemplateProcessor.GetTemplateFingerprint(dialog.FileName);
            _definitions.Clear();
            _definitions.AddRange(_templateProcessor.DiscoverPlaceholders(dialog.FileName));
            _slides = _templateProcessor.GetSlideSummaries(dialog.FileName);
            foreach (var definition in _definitions)
            {
                if (!_project.Values.ContainsKey(definition.Name))
                {
                    _project.Values[definition.Name] = definition.DefaultValue;
                }

                if (definition.Type == PlaceholderType.Image && !_imageScaling.ContainsKey(definition.Name))
                {
                    _imageScaling[definition.Name] = ImageScalingMode.Fit;
                }
            }

            BuildNavigation();
            BuildPropertyEditors();
            UpdatePreview();
            SetStatus($"Loaded {Path.GetFileName(dialog.FileName)} — {_definitions.Count} placeholders");
        }
        catch (Exception exception) when (exception is IOException or InvalidDataException or ArgumentException or OpenXmlPackageException)
        {
            ShowError("The selected template could not be loaded.", exception);
        }
    }

    private void BuildNavigation()
    {
        _slidesList.Items.Clear();
        foreach (var slide in _slides)
        {
            _slidesList.Items.Add($"Slide {slide.Number}");
        }

        if (_slidesList.Items.Count > 0)
        {
            _slidesList.SelectedIndex = 0;
        }

        _fieldsList.Items.Clear();
        foreach (var definition in _definitions)
        {
            _fieldsList.Items.Add(definition.DisplayName);
        }
    }

    private void BuildPropertyEditors()
    {
        _buildingControls = true;
        _valueReaders.Clear();
        _propertyEditor.SuspendLayout();
        _propertyEditor.Controls.Clear();
        foreach (var definition in _definitions)
        {
            var group = new Panel
            {
                Width = Math.Max(270, _propertyEditor.ClientSize.Width - 30),
                Height = definition.Type == PlaceholderType.MultilineText ? 106 : definition.Type == PlaceholderType.Image ? 126 : 78,
                Margin = new Padding(0, 0, 0, 9),
                BackColor = Color.FromArgb(248, 250, 252),
                Padding = new Padding(9)
            };
            var label = new Label
            {
                Dock = DockStyle.Top,
                Height = 22,
                Text = definition.DisplayName + (definition.Required ? " *" : string.Empty),
                Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold),
                ForeColor = Color.FromArgb(51, 65, 85)
            };
            group.Controls.Add(label);
            var currentValue = _project.Values.TryGetValue(definition.Name, out var value) ? value : definition.DefaultValue;
            var editor = BuildEditor(definition, currentValue, group);
            _valueReaders[definition.Name] = editor.ReadValue;
            _propertyEditor.Controls.Add(group);
        }

        _propertyEditor.ResumeLayout();
        _buildingControls = false;
        UpdatePreview();
    }

    private (Control Control, Func<string> ReadValue) BuildEditor(
        PlaceholderDefinition definition,
        string value,
        Panel group)
    {
        if (definition.Type == PlaceholderType.Image)
        {
            var pathBox = new TextBox
            {
                Dock = DockStyle.Top,
                ReadOnly = true,
                Text = value,
                PlaceholderText = "Choose an image..."
            };
            var buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 32,
                WrapContents = false,
                Padding = new Padding(0, 3, 0, 0)
            };
            var browse = new Button { Text = "Browse", AutoSize = true, Height = 25 };
            browse.Click += (_, _) =>
            {
                using var dialog = new OpenFileDialog
                {
                    Title = $"Select image for {definition.DisplayName}",
                    Filter = "Image files|*.png;*.jpg;*.jpeg;*.gif;*.bmp|All files|*.*"
                };
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    pathBox.Text = dialog.FileName;
                    UpdateValue(definition.Name, pathBox.Text);
                }
            };
            var clear = new Button { Text = "Clear", AutoSize = true, Height = 25 };
            clear.Click += (_, _) =>
            {
                pathBox.Clear();
                UpdateValue(definition.Name, string.Empty);
            };
            var scaling = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Width = 115,
                Height = 25,
                DataSource = Enum.GetValues<ImageScalingMode>()
            };
            var currentMode = _imageScaling.TryGetValue(definition.Name, out var mode) ? mode : ImageScalingMode.Fit;
            scaling.SelectedItem = currentMode;
            scaling.SelectedIndexChanged += (_, _) =>
            {
                if (scaling.SelectedItem is ImageScalingMode selected)
                {
                    _imageScaling[definition.Name] = selected;
                }
            };
            buttons.Controls.Add(browse);
            buttons.Controls.Add(clear);
            buttons.Controls.Add(scaling);
            group.Controls.Add(buttons);
            group.Controls.Add(pathBox);
            return (pathBox, () => pathBox.Text);
        }

        Control control;
        Func<string> readValue;
        switch (definition.Type)
        {
            case PlaceholderType.MultilineText:
                var multiLine = new TextBox { Dock = DockStyle.Fill, Multiline = true, ScrollBars = ScrollBars.Vertical, Text = value };
                multiLine.TextChanged += (_, _) => UpdateValue(definition.Name, multiLine.Text);
                control = multiLine;
                readValue = () => multiLine.Text;
                break;
            case PlaceholderType.Number:
                var numeric = new NumericUpDown
                {
                    Dock = DockStyle.Top,
                    DecimalPlaces = 2,
                    Minimum = -1000000000,
                    Maximum = 1000000000,
                    ThousandsSeparator = true
                };
                if (decimal.TryParse(value, out var number))
                {
                    numeric.Value = Math.Clamp(number, numeric.Minimum, numeric.Maximum);
                }

                numeric.ValueChanged += (_, _) => UpdateValue(definition.Name, numeric.Value.ToString());
                control = numeric;
                readValue = () => numeric.Value.ToString();
                break;
            case PlaceholderType.Date:
                var date = new DateTimePicker
                {
                    Dock = DockStyle.Top,
                    Format = DateTimePickerFormat.Short,
                    ShowCheckBox = true,
                    Checked = DateTime.TryParse(value, out var parsedDate)
                };
                if (date.Checked)
                {
                    date.Value = parsedDate;
                }

                date.ValueChanged += (_, _) => UpdateValue(definition.Name, date.Checked ? date.Value.ToString("d") : string.Empty);
                date.CloseUp += (_, _) => UpdateValue(definition.Name, date.Checked ? date.Value.ToString("d") : string.Empty);
                control = date;
                readValue = () => date.Checked ? date.Value.ToString("d") : string.Empty;
                break;
            case PlaceholderType.Boolean:
                var checkBox = new CheckBox { Dock = DockStyle.Top, Checked = bool.TryParse(value, out var checkedValue) && checkedValue };
                checkBox.CheckedChanged += (_, _) => UpdateValue(definition.Name, checkBox.Checked.ToString());
                control = checkBox;
                readValue = () => checkBox.Checked.ToString();
                break;
            case PlaceholderType.Dropdown:
                var combo = new ComboBox { Dock = DockStyle.Top, DropDownStyle = ComboBoxStyle.DropDown, Text = value };
                combo.TextChanged += (_, _) => UpdateValue(definition.Name, combo.Text);
                control = combo;
                readValue = () => combo.Text;
                break;
            default:
                var text = new TextBox { Dock = DockStyle.Top, Text = value };
                text.TextChanged += (_, _) => UpdateValue(definition.Name, text.Text);
                control = text;
                readValue = () => text.Text;
                break;
        }

        group.Controls.Add(control);
        return (control, readValue);
    }

    private void SelectSlide()
    {
        if (_slidesList.SelectedIndex < 0 || _slidesList.SelectedIndex >= _slides.Count)
        {
            return;
        }

        _selectedSlide = _slides[_slidesList.SelectedIndex].Number;
        UpdatePreview();
    }

    private void SelectFieldFromNavigation()
    {
        if (_fieldsList.SelectedIndex < 0 || _fieldsList.SelectedIndex >= _definitions.Count)
        {
            return;
        }

        var definition = _definitions[_fieldsList.SelectedIndex];
        var index = _slides.ToList().FindIndex(slide => slide.Number == definition.SlideNumber);
        if (index >= 0)
        {
            _navigationTabs.SelectedIndex = 0;
            _slidesList.SelectedIndex = index;
        }

        _navigationTabs.SelectedIndex = 1;
    }

    private void UpdateValue(string name, string value)
    {
        if (_buildingControls)
        {
            return;
        }

        _project.Values[name] = value;
        UpdatePreview();
    }

    private void UpdatePreview()
    {
        var slide = _slides.FirstOrDefault(item => item.Number == _selectedSlide);
        if (slide is null)
        {
            _preview.Text = string.IsNullOrWhiteSpace(_project.TemplatePath)
                ? "Load a .pptx template to begin. Placeholders such as {{REPORT_TITLE}} and {{IMAGE_TEST_SETUP}} are discovered automatically."
                : "This template contains no slides.";
            return;
        }

        var previewText = slide.Text;
        foreach (var definition in _definitions)
        {
            var value = _valueReaders.TryGetValue(definition.Name, out var reader)
                ? reader()
                : _project.Values.GetValueOrDefault(definition.Name, definition.DefaultValue);
            if (definition.Type == PlaceholderType.Image)
            {
                value = string.IsNullOrWhiteSpace(value) ? "[Select image]" : $"[Image: {Path.GetFileName(value)}]";
            }

            previewText = previewText.Replace("{{" + definition.Name + "}}", value, StringComparison.OrdinalIgnoreCase);
        }

        _preview.Text = $"SLIDE {_selectedSlide}{Environment.NewLine}{new string('-', 42)}{Environment.NewLine}{previewText}";
    }

    private void OpenReport()
    {
        using var dialog = new OpenFileDialog
        {
            Title = "Open a saved report",
            Filter = "Report projects (*.json)|*.json",
            CheckFileExists = true
        };
        if (dialog.ShowDialog(this) != DialogResult.OK)
        {
            return;
        }

        _ = OpenReportAsync(dialog.FileName);
    }

    private async Task OpenReportAsync(string path)
    {
        try
        {
            _project = await _projectStore.LoadAsync(path);
            _projectFilePath = path;
            _imageScaling.Clear();
            foreach (var item in _project.ImageScaling)
            {
                _imageScaling[item.Key] = item.Value;
            }

            if (!string.IsNullOrWhiteSpace(_project.TemplatePath) && File.Exists(_project.TemplatePath))
            {
                var fingerprint = OpenXmlTemplateProcessor.GetTemplateFingerprint(_project.TemplatePath);
                if (!string.Equals(fingerprint, _project.TemplateFingerprint, StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show(this, "The source template has changed since this report was saved.", "Template changed",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                _definitions.Clear();
                _definitions.AddRange(_templateProcessor.DiscoverPlaceholders(_project.TemplatePath));
                _slides = _templateProcessor.GetSlideSummaries(_project.TemplatePath);
                BuildNavigation();
                BuildPropertyEditors();
                SetStatus($"Opened {Path.GetFileName(path)}");
            }
            else
            {
                _definitions.Clear();
                _slides = [];
                BuildNavigation();
                BuildPropertyEditors();
                ShowError("The report's PowerPoint template could not be found.",
                    new FileNotFoundException("Template path is missing or no longer exists.", _project.TemplatePath));
            }

            UpdatePreview();
        }
        catch (Exception exception) when (exception is IOException or InvalidDataException or JsonException or ArgumentException or OpenXmlPackageException)
        {
            ShowError("The report project could not be opened.", exception);
        }
    }

    private void SaveReport()
    {
        if (_projectFilePath is null)
        {
            SaveReportAs();
            return;
        }

        _ = SaveReportAsync(_projectFilePath);
    }

    private void SaveReportAs()
    {
        using var dialog = new SaveFileDialog
        {
            Title = "Save report project",
            Filter = "Report projects (*.json)|*.json",
            DefaultExt = "json",
            FileName = Path.GetInvalidFileNameChars().Aggregate(
                _project.ReportName,
                (name, invalidCharacter) => name.Replace(invalidCharacter, '_'))
        };
        if (dialog.ShowDialog(this) == DialogResult.OK)
        {
            _projectFilePath = dialog.FileName;
            _project.ReportName = Path.GetFileNameWithoutExtension(dialog.FileName);
            _ = SaveReportAsync(dialog.FileName);
        }
    }

    private async Task SaveReportAsync(string path)
    {
        try
        {
            ReadCurrentValues();
            _project.ImageScaling = new Dictionary<string, ImageScalingMode>(_imageScaling, StringComparer.OrdinalIgnoreCase);
            await _projectStore.SaveAsync(path, _project);
            SetStatus($"Saved {Path.GetFileName(path)}");
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or ArgumentException)
        {
            ShowError("The report project could not be saved.", exception);
        }
    }

    private async void GeneratePresentation()
    {
        if (string.IsNullOrWhiteSpace(_project.TemplatePath) || !File.Exists(_project.TemplatePath))
        {
            MessageBox.Show(this, "Load a PowerPoint template before generating a report.", "Template required",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        ReadCurrentValues();
        var issues = ReportValidation.Validate(_definitions, _project.Values);
        if (issues.Count > 0)
        {
            var message = string.Join(Environment.NewLine, issues.Select(issue => $"- {issue.Message}"));
            MessageBox.Show(this, message, "Please fix these fields", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        using var dialog = new SaveFileDialog
        {
            Title = "Generate PowerPoint report",
            Filter = "PowerPoint presentations (*.pptx)|*.pptx",
            DefaultExt = "pptx",
            FileName = $"{_project.ReportName}.pptx"
        };
        if (dialog.ShowDialog(this) != DialogResult.OK)
        {
            return;
        }

        try
        {
            SetStatus("Generating presentation...");
            await _templateProcessor.GenerateAsync(_project.TemplatePath, dialog.FileName, _project.Values, _imageScaling);
            SetStatus($"Generated {Path.GetFileName(dialog.FileName)}");
            MessageBox.Show(this, "The report was generated successfully.", "Report complete",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception exception) when (exception is IOException or InvalidDataException or ArgumentException or OpenXmlPackageException)
        {
            SetStatus("Generation failed");
            ShowError("The PowerPoint report could not be generated.", exception);
        }
    }

    private void ReadCurrentValues()
    {
        foreach (var (name, readValue) in _valueReaders)
        {
            _project.Values[name] = readValue();
        }
    }

    private void ShowError(string message, Exception exception) =>
        MessageBox.Show(this, $"{message}{Environment.NewLine}{exception.Message}", "PowerPoint Report Generator",
            MessageBoxButtons.OK, MessageBoxIcon.Error);

    private void SetStatus(string message) => _statusLabel.Text = message;
}
