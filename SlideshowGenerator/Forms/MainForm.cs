using LanguageCourseSlides.Services;
using SlideshowGenerator.Models;
using SlideshowGenerator.Services;

namespace SlideshowGenerator.Forms;

public class MainForm : Form
{
    // ── state ────────────────────────────────────────────────────────────
    private List<WordEntry>    _entries       = [];
    private TemplateConfig?    _activeConfig;
    private string?            _excelPath;

    // ── controls ─────────────────────────────────────────────────────────
    private ComboBox     cmbTemplates    = null!;
    private Button       btnManage       = null!;
    private Label        lblTemplateSummary = null!;
    private TextBox      txtExcelPath    = null!;
    private Button       btnLoadExcel    = null!;
    private Label        lblEntryCount   = null!;
    private DataGridView dgvPreview      = null!;
    private Button       btnGenerate     = null!;
    private ProgressBar  progressBar     = null!;
    private Label        lblStatus       = null!;

    // column indices for asset columns
    private DataGridViewTextBoxColumn colImage = null!;
    private DataGridViewTextBoxColumn colAudio = null!;

    // ── constructor ──────────────────────────────────────────────────────

    public MainForm()
    {
        BuildUI();
        RefreshTemplateList();
    }

    // ── UI builder ───────────────────────────────────────────────────────

    private void BuildUI()
    {
        Text            = "Language Course Slide Generator";
        Size            = new Size(1000, 680);
        MinimumSize     = new Size(800, 550);
        StartPosition   = FormStartPosition.CenterScreen;
        Font            = new Font("Segoe UI", 9f);

        // ─────────────────── top control strip ───────────────────────────
        var topPanel = new TableLayoutPanel
        {
            Dock        = DockStyle.Top,
            Height      = 105,
            ColumnCount = 3,
            RowCount    = 3,
            Padding     = new Padding(10, 10, 10, 4),
        };
        topPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));
        topPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,  100));
        topPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100));

        // Row 0 — template selector
        topPanel.Controls.Add(MakeLabel("Template:"), 0, 0);

        cmbTemplates = new ComboBox
        {
            Dock          = DockStyle.Fill,
            DropDownStyle = ComboBoxStyle.DropDownList,
            DisplayMember = nameof(TemplateConfig.ConfigName),
            Margin        = new Padding(0, 2, 4, 4),
        };
        cmbTemplates.SelectedIndexChanged += CmbTemplates_Changed;
        topPanel.Controls.Add(cmbTemplates, 1, 0);

        btnManage = new Button
        {
            Text   = "Manage…",
            Dock   = DockStyle.Fill,
            Margin = new Padding(0, 2, 0, 4),
        };
        btnManage.Click += (s, e) =>
        {
            using var mgr = new TemplateManagerForm();
            mgr.ShowDialog(this);
            RefreshTemplateList();
        };
        topPanel.Controls.Add(btnManage, 2, 0);

        // Row 1 — template summary
        lblTemplateSummary = new Label
        {
            Dock      = DockStyle.Fill,
            ForeColor = Color.Gray,
            Margin    = new Padding(0, 0, 0, 4),
        };
        topPanel.Controls.Add(lblTemplateSummary, 1, 1);
        topPanel.SetColumnSpan(lblTemplateSummary, 2);

        // Row 2 — excel picker
        topPanel.Controls.Add(MakeLabel("Word List (.xlsx):"), 0, 2);

        txtExcelPath = new TextBox
        {
            Dock      = DockStyle.Fill,
            ReadOnly  = true,
            Margin    = new Padding(0, 2, 4, 0),
            BackColor = SystemColors.Window,
        };
        topPanel.Controls.Add(txtExcelPath, 1, 2);

        btnLoadExcel = new Button
        {
            Text   = "Browse…",
            Dock   = DockStyle.Fill,
            Margin = new Padding(0, 2, 0, 0),
        };
        btnLoadExcel.Click += BtnLoadExcel_Click;
        topPanel.Controls.Add(btnLoadExcel, 2, 2);

        // ─────────────────── entry count label ───────────────────────────
        lblEntryCount = new Label
        {
            Dock      = DockStyle.Top,
            Height    = 22,
            Padding   = new Padding(10, 2, 0, 0),
            ForeColor = Color.DimGray,
            Text      = "No word list loaded.",
        };

        // ─────────────────── preview grid ────────────────────────────────
        dgvPreview = new DataGridView
        {
            Dock                  = DockStyle.Fill,
            AllowUserToAddRows    = false,
            AllowUserToDeleteRows = false,
            ReadOnly              = true,
            RowHeadersVisible     = false,
            SelectionMode         = DataGridViewSelectionMode.FullRowSelect,
            BackgroundColor       = SystemColors.Window,
            BorderStyle           = BorderStyle.None,
            AutoSizeColumnsMode   = DataGridViewAutoSizeColumnsMode.None,
            ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
            ColumnHeadersHeight   = 26,
        };

        // Define preview columns
        AddTextCol("Word",          "Word",          80);
        AddTextCol("Plural",        "Plural",        70);
        AddTextCol("Type",          "Type",          60);
        AddTextCol("Stress",        "Stress",        55);
        AddTextCol("Vowels",        "Vowels",        55);
        AddTextCol("Hint",          "Hint",          90);
        AddTextCol("Trap",          "Trap",          90);
        AddTextCol("Rule",          "Rule",          90);
        AddTextCol("Usage",         "Usage",         90);
        AddTextCol("English",       "English",       90);
        AddTextCol("Pronunciation", "Pronunciation", 110);

        colImage = AddTextCol("ImageStatus", "Image", 120);
        colAudio = AddTextCol("AudioStatus", "Audio", 120);

        dgvPreview.CellClick         += DgvPreview_CellClick;
        dgvPreview.CellFormatting    += DgvPreview_CellFormatting;

        // ─────────────────── bottom strip ────────────────────────────────
        var bottomPanel = new Panel
        {
            Dock   = DockStyle.Bottom,
            Height = 50,
        };

        progressBar = new ProgressBar
        {
            Location = new Point(10, 14),
            Width    = 300,
            Height   = 22,
            Visible  = false,
        };

        btnGenerate = new Button
        {
            Text     = "Generate Presentation",
            Size     = new Size(180, 30),
            Location = new Point(bottomPanel.Width - 198, 10),
            Anchor   = AnchorStyles.Right | AnchorStyles.Top,
            Enabled  = false,
            BackColor = Color.FromArgb(0, 120, 215),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font      = new Font("Segoe UI", 9.5f, FontStyle.Bold),
        };
        btnGenerate.FlatAppearance.BorderSize = 0;
        btnGenerate.Click += BtnGenerate_Click;

        lblStatus = new Label
        {
            Location  = new Point(320, 17),
            AutoSize  = true,
            ForeColor = Color.DimGray,
        };

        bottomPanel.Controls.AddRange([progressBar, btnGenerate, lblStatus]);

        // ─────────────────── assemble ─────────────────────────────────────
        Controls.Add(dgvPreview);
        Controls.Add(lblEntryCount);
        Controls.Add(topPanel);
        Controls.Add(bottomPanel);
    }

    private DataGridViewTextBoxColumn AddTextCol(string prop, string header, int width)
    {
        var col = new DataGridViewTextBoxColumn
        {
            DataPropertyName = prop,
            HeaderText       = header,
            Width            = width,
            SortMode         = DataGridViewColumnSortMode.NotSortable,
        };
        dgvPreview.Columns.Add(col);
        return col;
    }

    // ── event handlers ───────────────────────────────────────────────────

    private void CmbTemplates_Changed(object? sender, EventArgs e)
    {
        _activeConfig = cmbTemplates.SelectedItem as TemplateConfig;
        UpdateTemplateSummary();
        UpdateGenerateButton();
    }

    private void UpdateTemplateSummary()
    {
        if (_activeConfig == null)
        {
            lblTemplateSummary.Text = "";
            return;
        }
        var c = _activeConfig;
        lblTemplateSummary.Text =
            $"Slides: {c.StaticSlides.Count} static, " +
            $"{(c.IndexSlide != null ? 1 : 0)} index, " +
            $"{(c.WordSlide  != null ? 1 : 0)} word  |  " +
            $"Hyperlinks: {(c.HyperlinkIndex ? "Yes" : "No")}  |  " +
            $"File: {Path.GetFileName(c.TemplatePath)}";
    }

    private void BtnLoadExcel_Click(object? sender, EventArgs e)
    {
        using var dlg = new OpenFileDialog
        {
            Title  = "Select Word List Spreadsheet",
            Filter = "Excel Files|*.xlsx",
        };
        if (dlg.ShowDialog() != DialogResult.OK) return;

        _excelPath = dlg.FileName;
        txtExcelPath.Text = dlg.FileName;
        SetStatus("Reading spreadsheet…");

        try
        {
            _entries = ExcelReader.Read(_excelPath);
            BindPreview();
            lblEntryCount.Text =
                $"{_entries.Count} word(s) loaded  —  " +
                $"{_entries.Count(x => x.HasImage)} with images, " +
                $"{_entries.Count(x => x.HasAudio)} with audio.";
            SetStatus("Ready.");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to read spreadsheet:\n\n{ex.Message}",
                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            SetStatus("Load failed.");
        }

        UpdateGenerateButton();
    }

    private void BindPreview()
    {
        dgvPreview.AutoGenerateColumns = false;
        dgvPreview.DataSource = null;
        dgvPreview.DataSource = _entries;
    }

    // Clicking an asset cell lets the user browse to override the path
    private void DgvPreview_CellClick(object? sender, DataGridViewCellEventArgs e)
    {
        if (e.RowIndex < 0) return;

        bool isImg = e.ColumnIndex == colImage.Index;
        bool isAud = e.ColumnIndex == colAudio.Index;
        if (!isImg && !isAud) return;

        var entry = _entries[e.RowIndex];

        using var dlg = new OpenFileDialog();
        if (isImg)
        {
            dlg.Title  = "Select Image for this word";
            dlg.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.gif;*.webp;*.bmp";
        }
        else
        {
            dlg.Title  = "Select Audio for this word";
            dlg.Filter = "Audio Files|*.mp3;*.wav;*.m4a;*.ogg";
        }

        if (dlg.ShowDialog() != DialogResult.OK) return;

        var asset = AssetData.FromPath(dlg.FileName, Path.GetDirectoryName(dlg.FileName)!);
        if (isImg) entry.Image = asset;
        else       entry.Audio = asset;

        // Refresh just this row
        dgvPreview.InvalidateRow(e.RowIndex);
        UpdateGenerateButton();
    }

    // Colour-code asset cells green/grey
    private void DgvPreview_CellFormatting(object? sender, DataGridViewCellFormattingEventArgs e)
    {
        if (e.RowIndex < 0 || e.RowIndex >= _entries.Count) return;

        var entry = _entries[e.RowIndex];

        if (e.ColumnIndex == colImage.Index)
        {
            e.Value = entry.ImageStatus;
            e.CellStyle.BackColor = entry.HasImage
                ? Color.FromArgb(220, 255, 220)
                : SystemColors.Window;
            e.FormattingApplied = true;
        }
        else if (e.ColumnIndex == colAudio.Index)
        {
            e.Value = entry.AudioStatus;
            e.CellStyle.BackColor = entry.HasAudio
                ? Color.FromArgb(210, 235, 255)
                : SystemColors.Window;
            e.FormattingApplied = true;
        }
    }

    private async void BtnGenerate_Click(object? sender, EventArgs e)
    {
        if (_activeConfig == null || _entries.Count == 0) return;

        using var dlg = new SaveFileDialog
        {
            Title    = "Save Generated Presentation",
            Filter   = "PowerPoint|*.pptx",
            FileName = $"Slides_{DateTime.Now:yyyyMMdd_HHmm}.pptx",
        };
        if (dlg.ShowDialog() != DialogResult.OK) return;

        SetGenerating(true);

        var progress = new Progress<int>(done =>
        {
            progressBar.Value = Math.Min(done, progressBar.Maximum);
            lblStatus.Text    = $"Generating slide {done} of {_entries.Count}…";
        });

        try
        {
            var config  = _activeConfig;
            var entries = _entries.ToList();
            var output  = dlg.FileName;

            await Task.Run(() => PptxGenerator.Generate(config, output, entries, progress));

            SetStatus($"Done — {_entries.Count} slide(s) generated.");

            if (MessageBox.Show(
                    $"Presentation created successfully!\n\n{output}\n\nOpen it now?",
                    "Success",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Information) == DialogResult.Yes)
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName        = output,
                    UseShellExecute = true
                });
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Generation failed:\n\n{ex.Message}",
                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            SetStatus("Generation failed.");
        }
        finally
        {
            SetGenerating(false);
        }
    }

    // ── helpers ──────────────────────────────────────────────────────────

    private void RefreshTemplateList()
    {
        var configs  = ConfigManager.LoadAll();
        var previous = _activeConfig?.ConfigName;

        cmbTemplates.DataSource = null;
        cmbTemplates.DataSource = configs;

        // Restore previous selection if it still exists
        if (previous != null)
        {
            var match = configs.FirstOrDefault(c => c.ConfigName == previous);
            if (match != null) cmbTemplates.SelectedItem = match;
        }

        if (cmbTemplates.SelectedItem == null && configs.Count > 0)
            cmbTemplates.SelectedIndex = 0;

        _activeConfig = cmbTemplates.SelectedItem as TemplateConfig;
        UpdateTemplateSummary();
        UpdateGenerateButton();
    }

    private void UpdateGenerateButton()
    {
        btnGenerate.Enabled =
            _activeConfig != null      &&
            _activeConfig.IsValid      &&
            _entries.Count > 0;
    }

    private void SetGenerating(bool generating)
    {
        btnGenerate.Enabled   = !generating;
        btnLoadExcel.Enabled  = !generating;
        btnManage.Enabled     = !generating;
        cmbTemplates.Enabled  = !generating;
        progressBar.Visible   = generating;
        progressBar.Maximum   = Math.Max(1, _entries.Count);
        progressBar.Value     = 0;
    }

    private void SetStatus(string msg) => lblStatus.Text = msg;

    private static Label MakeLabel(string text) =>
        new() { Text = text, AutoSize = true,
                Anchor = AnchorStyles.Left | AnchorStyles.Top,
                Margin = new Padding(0, 5, 8, 0) };
}
