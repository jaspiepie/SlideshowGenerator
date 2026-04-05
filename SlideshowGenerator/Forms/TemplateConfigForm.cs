using DocumentFormat.OpenXml.Packaging;
using SlideshowGenerator.Models;
using SlideshowGenerator.Services;

namespace SlideshowGenerator.Forms;

/// <summary>
/// Create or edit a TemplateConfig. Shows all slides found in the chosen
/// .pptx and lets the user assign a Role to each one.
/// </summary>
public class TemplateConfigForm : Form
{
    // ── state ────────────────────────────────────────────────────────────
    private TemplateConfig _config;
    private readonly bool  _isEdit;
    private string?        _originalName;

    public TemplateConfig? Result { get; private set; }

    // ── controls ─────────────────────────────────────────────────────────
    private TextBox         txtName           = null!;
    private TextBox         txtTemplatePath   = null!;
    private Button          btnBrowseTemplate = null!;
    private DataGridView    dgvSlides         = null!;
    private Label           lblSlideInfo      = null!;
    private TextBox         txtIndexFormat    = null!;
    private TextBox         txtImagePH        = null!;
    private TextBox         txtAudioPH        = null!;
    private CheckBox        chkHyperlink      = null!;
    private Button          btnSave           = null!;
    private Button          btnCancel         = null!;

    // ── constructor ──────────────────────────────────────────────────────

    public TemplateConfigForm(TemplateConfig? existing = null)
    {
        _isEdit = existing != null;
        _config = existing != null ? ConfigManager.Clone(existing) : new TemplateConfig();
        _originalName = existing?.ConfigName;

        BuildUI();

        if (_isEdit) PopulateFromConfig();
    }

    // ── UI builder ───────────────────────────────────────────────────────

    private void BuildUI()
    {
        Text            = _isEdit ? "Edit Template" : "New Template";
        Size            = new Size(700, 580);
        MinimumSize     = new Size(600, 500);
        StartPosition   = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.Sizable;
        Font            = new Font("Segoe UI", 9f);

        // ── top panel ────────────────────────────────────────────────────
        var topPanel = new TableLayoutPanel
        {
            Dock        = DockStyle.Top,
            Height      = 130,
            ColumnCount = 3,
            RowCount    = 3,
            Padding     = new Padding(10, 10, 10, 4),
        };
        topPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 130));
        topPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,  100));
        topPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));

        topPanel.Controls.Add(MakeLabel("Template Name:"), 0, 0);
        txtName = new TextBox { Dock = DockStyle.Fill, Margin = new Padding(0, 2, 4, 4) };
        topPanel.Controls.Add(txtName, 1, 0);
        topPanel.SetColumnSpan(txtName, 2);

        topPanel.Controls.Add(MakeLabel("Template File (.pptx):"), 0, 1);
        txtTemplatePath = new TextBox
        {
            Dock     = DockStyle.Fill,
            ReadOnly = true,
            Margin   = new Padding(0, 2, 4, 4),
            BackColor = SystemColors.Window
        };
        topPanel.Controls.Add(txtTemplatePath, 1, 1);

        btnBrowseTemplate = new Button { Text = "Browse…", Dock = DockStyle.Fill,
                                         Margin = new Padding(0, 2, 0, 4) };
        btnBrowseTemplate.Click += BrowseTemplate_Click;
        topPanel.Controls.Add(btnBrowseTemplate, 2, 1);

        lblSlideInfo = new Label
        {
            Text      = "Load a .pptx to see its slides.",
            Dock      = DockStyle.Fill,
            ForeColor = Color.Gray,
            Margin    = new Padding(0, 2, 0, 0)
        };
        topPanel.Controls.Add(lblSlideInfo, 1, 2);
        topPanel.SetColumnSpan(lblSlideInfo, 2);

        // ── slide grid ───────────────────────────────────────────────────
        dgvSlides = new DataGridView
        {
            Dock                  = DockStyle.Fill,
            AllowUserToAddRows    = false,
            AllowUserToDeleteRows = false,
            RowHeadersVisible     = false,
            SelectionMode         = DataGridViewSelectionMode.FullRowSelect,
            BackgroundColor       = SystemColors.Window,
            BorderStyle           = BorderStyle.None,
            AutoSizeRowsMode      = DataGridViewAutoSizeRowsMode.AllCells,
        };

        dgvSlides.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "colNum", HeaderText = "#",
            ReadOnly = true, Width = 40, Resizable = DataGridViewTriState.False
        });
        dgvSlides.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "colLabel", HeaderText = "Slide Label", Width = 160
        });

        var roleCol = new DataGridViewComboBoxColumn
        {
            Name = "colRole", HeaderText = "Role", Width = 90,
            FlatStyle = FlatStyle.Flat
        };
        roleCol.Items.AddRange("Static", "Index", "Word");
        dgvSlides.Columns.Add(roleCol);

        dgvSlides.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "colPlaceholder", HeaderText = "Index Placeholder",
            Width = 160,
            ToolTipText = "Only used when Role = Index"
        });

        dgvSlides.CellValueChanged     += DgvSlides_CellValueChanged;
        dgvSlides.CurrentCellDirtyStateChanged += (s, e) =>
        {
            if (dgvSlides.IsCurrentCellDirty) dgvSlides.CommitEdit(DataGridViewDataErrorContexts.Commit);
        };

        // ── options panel ────────────────────────────────────────────────
        var optPanel = new TableLayoutPanel
        {
            Dock        = DockStyle.Bottom,
            Height      = 110,
            ColumnCount = 4,
            RowCount    = 3,
            Padding     = new Padding(10, 4, 10, 4),
        };
        optPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 130));
        optPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,  50));
        optPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 130));
        optPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,  50));

        optPanel.Controls.Add(MakeLabel("Index Line Format:"), 0, 0);
        txtIndexFormat = new TextBox
        {
            Dock = DockStyle.Fill,
            Text = "{n}. {word}",
            Margin = new Padding(0, 2, 8, 4)
        };
        optPanel.Controls.Add(txtIndexFormat, 1, 0);

        var fmtHint = new Label
        {
            Text      = "Tokens: {n}  {word}  {plural}  {type}  {english}",
            Dock      = DockStyle.Fill,
            ForeColor = Color.Gray,
            Margin    = new Padding(0, 4, 0, 0)
        };
        optPanel.Controls.Add(fmtHint, 2, 0);
        optPanel.SetColumnSpan(fmtHint, 2);

        optPanel.Controls.Add(MakeLabel("Image Placeholder:"), 0, 1);
        txtImagePH = new TextBox { Dock = DockStyle.Fill, Text = "{{Image}}",
                                   Margin = new Padding(0, 2, 8, 4) };
        optPanel.Controls.Add(txtImagePH, 1, 1);

        optPanel.Controls.Add(MakeLabel("Audio Placeholder:"), 2, 1);
        txtAudioPH = new TextBox { Dock = DockStyle.Fill, Text = "{{Audio}}",
                                   Margin = new Padding(0, 2, 0, 4) };
        optPanel.Controls.Add(txtAudioPH, 3, 1);

        chkHyperlink = new CheckBox
        {
            Text    = "Hyperlink index entries to word slides",
            Checked = true,
            Dock    = DockStyle.Fill,
            Margin  = new Padding(0, 2, 0, 0)
        };
        optPanel.Controls.Add(chkHyperlink, 0, 2);
        optPanel.SetColumnSpan(chkHyperlink, 4);

        // ── button row ───────────────────────────────────────────────────
        var btnPanel = new FlowLayoutPanel
        {
            Dock          = DockStyle.Bottom,
            Height        = 40,
            FlowDirection = FlowDirection.RightToLeft,
            Padding       = new Padding(8, 4, 8, 4),
        };

        btnCancel = new Button { Text = "Cancel", Width = 80,
                                  DialogResult = DialogResult.Cancel };
        btnSave   = new Button { Text = "Save",   Width = 80 };
        btnSave.Click += BtnSave_Click;

        btnPanel.Controls.AddRange([btnCancel, btnSave]);

        Controls.Add(dgvSlides);
        Controls.Add(topPanel);
        Controls.Add(optPanel);
        Controls.Add(btnPanel);
    }

    // ── event handlers ───────────────────────────────────────────────────

    private void BrowseTemplate_Click(object? sender, EventArgs e)
    {
        using var dlg = new OpenFileDialog
        {
            Title  = "Select PowerPoint Template",
            Filter = "PowerPoint Files|*.pptx"
        };
        if (dlg.ShowDialog() != DialogResult.OK) return;

        _config.TemplatePath = dlg.FileName;
        txtTemplatePath.Text = dlg.FileName;
        LoadSlidesFromPptx(dlg.FileName);
    }

    private void LoadSlidesFromPptx(string path)
    {
        dgvSlides.Rows.Clear();
        try
        {
            using var prs   = PresentationDocument.Open(path, isEditable: false);
            int slideCount  = prs.PresentationPart!.SlideParts.Count();

            for (int i = 0; i < slideCount; i++)
            {
                var existing = _config.Slides.FirstOrDefault(s => s.SlideIndex == i);
                var row      = dgvSlides.Rows.Add();
                var r        = dgvSlides.Rows[row];

                r.Cells["colNum"].Value         = i + 1;
                r.Cells["colLabel"].Value       = existing?.Label ?? $"Slide {i + 1}";
                r.Cells["colRole"].Value        = existing?.Role.ToString() ?? "Static";
                r.Cells["colPlaceholder"].Value = existing?.IndexPlaceholder ?? "{{Index}}";
            }

            lblSlideInfo.Text      = $"{slideCount} slide(s) found in template.";
            lblSlideInfo.ForeColor = Color.DarkGreen;
            UpdatePlaceholderColumnVisibility();
        }
        catch (Exception ex)
        {
            lblSlideInfo.Text      = $"Could not read template: {ex.Message}";
            lblSlideInfo.ForeColor = Color.Red;
        }
    }

    private void DgvSlides_CellValueChanged(object? sender, DataGridViewCellEventArgs e)
    {
        if (e.ColumnIndex == dgvSlides.Columns["colRole"]!.Index)
            UpdatePlaceholderColumnVisibility();
    }

    private void UpdatePlaceholderColumnVisibility()
    {
        // Dim the placeholder cell for non-Index rows
        foreach (DataGridViewRow row in dgvSlides.Rows)
        {
            var role = row.Cells["colRole"].Value?.ToString();
            var cell = row.Cells["colPlaceholder"];
            cell.Style.BackColor = role == "Index"
                ? SystemColors.Window
                : Color.FromArgb(240, 240, 240);
            cell.ReadOnly = role != "Index";
        }
    }

    private void BtnSave_Click(object? sender, EventArgs e)
    {
        if (string.IsNullOrWhiteSpace(txtName.Text))
        {
            ShowError("Please enter a template name.");
            return;
        }
        if (string.IsNullOrWhiteSpace(_config.TemplatePath) ||
            !File.Exists(_config.TemplatePath))
        {
            ShowError("Please select a valid .pptx template file.");
            return;
        }
        if (dgvSlides.Rows.Count == 0)
        {
            ShowError("Load a template file first.");
            return;
        }

        var slides = CollectSlides();

        if (!slides.Any(s => s.Role == SlideRole.Word))
        {
            ShowError("At least one slide must be assigned the 'Word' role.");
            return;
        }
        if (slides.Count(s => s.Role == SlideRole.Word) > 1)
        {
            ShowError("Only one slide can be the 'Word' slide.");
            return;
        }

        _config.ConfigName       = txtName.Text.Trim();
        _config.Slides           = slides;
        _config.IndexLineFormat  = txtIndexFormat.Text.Trim().IfEmpty("{n}. {word}");
        _config.ImagePlaceholder = txtImagePH.Text.Trim().IfEmpty("{{Image}}");
        _config.AudioPlaceholder = txtAudioPH.Text.Trim().IfEmpty("{{Audio}}");
        _config.HyperlinkIndex   = chkHyperlink.Checked;

        if (_isEdit && _originalName != null && _originalName != _config.ConfigName)
            ConfigManager.Rename(_config, _originalName);
        else
            ConfigManager.Save(_config);

        Result = _config;
        DialogResult = DialogResult.OK;
        Close();
    }

    // ── helpers ──────────────────────────────────────────────────────────

    private List<SlideDefinition> CollectSlides()
    {
        var list = new List<SlideDefinition>();
        foreach (DataGridViewRow row in dgvSlides.Rows)
        {
            if (row.IsNewRow) continue;
            Enum.TryParse<SlideRole>(row.Cells["colRole"].Value?.ToString(), out var role);
            list.Add(new SlideDefinition
            {
                SlideIndex       = (int)row.Cells["colNum"].Value! - 1,
                Label            = row.Cells["colLabel"].Value?.ToString() ?? "",
                Role             = role,
                IndexPlaceholder = row.Cells["colPlaceholder"].Value?.ToString() ?? "{{Index}}"
            });
        }
        return list.OrderBy(s => s.SlideIndex).ToList();
    }

    private void PopulateFromConfig()
    {
        txtName.Text          = _config.ConfigName;
        txtTemplatePath.Text  = _config.TemplatePath;
        txtIndexFormat.Text   = _config.IndexLineFormat;
        txtImagePH.Text       = _config.ImagePlaceholder;
        txtAudioPH.Text       = _config.AudioPlaceholder;
        chkHyperlink.Checked  = _config.HyperlinkIndex;

        if (File.Exists(_config.TemplatePath))
            LoadSlidesFromPptx(_config.TemplatePath);
    }

    private static Label MakeLabel(string text) =>
        new() { Text = text, AutoSize = true,
                Anchor = AnchorStyles.Left | AnchorStyles.Top,
                Margin = new Padding(0, 5, 8, 0) };

    private void ShowError(string msg) =>
        MessageBox.Show(msg, "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
}

// small string extension used only here
file static class StringExtensions
{
    public static string IfEmpty(this string s, string fallback) =>
        string.IsNullOrWhiteSpace(s) ? fallback : s;
}
