using LanguageCourseSlides.Models;
using LanguageCourseSlides.Services;
using LanguageCourseSlides.Forms;

namespace LanguageCourseSlides.Forms;

public class MainForm : Form
{
    // State
    private List<WordEntry> _entries      = [];
    private TemplateConfig? _activeConfig;
    private string?         _excelPath;

    // Controls
    private ComboBox     cmbTemplates       = null!;
    private Button       btnManage          = null!;
    private Label        lblTemplateSummary = null!;
    private TextBox      txtExcelPath       = null!;
    private Button       btnLoadExcel       = null!;
    private Label        lblEntryCount      = null!;
    private DataGridView dgvPreview         = null!;
    private Button       btnGenerate        = null!;
    private ProgressBar  progressBar        = null!;
    private Label        lblStatus          = null!;

    private DataGridViewTextBoxColumn colImage = null!;
    private DataGridViewTextBoxColumn colAudio = null!;

    public MainForm()
    {
        BuildUI();
        RefreshTemplateList();
    }

    // ── UI construction ──────────────────────────────────────────────────

    private void BuildUI()
    {
        Text          = "Language Course Slide Generator";
        Size          = new Size(1040, 680);
        MinimumSize   = new Size(800, 520);
        StartPosition = FormStartPosition.CenterScreen;
        Font          = new Font("Segoe UI", 9f);

        // ── Top strip ────────────────────────────────────────────────────
        var top = new TableLayoutPanel
        {
            Dock = DockStyle.Top, Height = 110, ColumnCount = 3, RowCount = 3,
            Padding = new Padding(10, 10, 10, 4),
        };
        top.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 130));
        top.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        top.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 108));

        // Row 0 — template
        top.Controls.Add(MkLabel("Template:"), 0, 0);
        cmbTemplates = new ComboBox
        {
            Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList,
            DisplayMember = nameof(TemplateConfig.ConfigName), Margin = new Padding(0,2,4,4),
        };
        cmbTemplates.SelectedIndexChanged += CmbTemplates_Changed;
        top.Controls.Add(cmbTemplates, 1, 0);
        btnManage = new Button { Text = "Manage…", Dock = DockStyle.Fill, Margin = new Padding(0,2,0,4) };
        btnManage.Click += (_, _) => { using var f = new TemplateManagerForm(); f.ShowDialog(this); RefreshTemplateList(); };
        top.Controls.Add(btnManage, 2, 0);

        // Row 1 — template summary
        lblTemplateSummary = new Label { Dock = DockStyle.Fill, ForeColor = Color.Gray, Margin = new Padding(0,0,0,4) };
        top.Controls.Add(lblTemplateSummary, 1, 1);
        top.SetColumnSpan(lblTemplateSummary, 2);

        // Row 2 — excel
        top.Controls.Add(MkLabel("Word List (.xlsx):"), 0, 2);
        txtExcelPath = new TextBox { Dock = DockStyle.Fill, ReadOnly = true, Margin = new Padding(0,2,4,0), BackColor = SystemColors.Window };
        top.Controls.Add(txtExcelPath, 1, 2);
        btnLoadExcel = new Button { Text = "Browse…", Dock = DockStyle.Fill, Margin = new Padding(0,2,0,0) };
        btnLoadExcel.Click += LoadExcel;
        top.Controls.Add(btnLoadExcel, 2, 2);

        // ── Entry count ───────────────────────────────────────────────────
        lblEntryCount = new Label
        {
            Dock = DockStyle.Top, Height = 22,
            Padding = new Padding(10, 2, 0, 0), ForeColor = Color.DimGray,
            Text = "No word list loaded.",
        };

        // ── Grid ──────────────────────────────────────────────────────────
        dgvPreview = new DataGridView
        {
            Dock = DockStyle.Fill, AllowUserToAddRows = false, AllowUserToDeleteRows = false,
            ReadOnly = true, RowHeadersVisible = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            BackgroundColor = SystemColors.Window, BorderStyle = BorderStyle.None,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
            ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
            ColumnHeadersHeight = 26,
        };

        // Text columns — bound by DataPropertyName
        AddCol("Word",          "Word",          80);
        AddCol("Plural",        "Plural",        70);
        AddCol("Type",          "Type",          60);
        AddCol("Stress",        "Stress",        56);
        AddCol("Vowels",        "Vowels",        56);
        AddCol("Hint",          "Hint",          90);
        AddCol("Trap",          "Trap",          90);
        AddCol("Rule",          "Rule",          90);
        AddCol("Usage",         "Usage",         90);
        AddCol("English",       "English",       90);
        AddCol("Pronunciation", "Pronunciation", 110);

        colImage = AddCol("ImageStatus", "Image 🖼", 120);
        colAudio = AddCol("AudioStatus", "Audio 🔊", 120);

        dgvPreview.CellClick      += Grid_CellClick;
        dgvPreview.CellFormatting += Grid_CellFormatting;

        // ── Bottom bar ────────────────────────────────────────────────────
        var bottom = new Panel { Dock = DockStyle.Bottom, Height = 54 };

        progressBar = new ProgressBar { Location = new Point(10, 16), Width = 280, Height = 22, Visible = false };

        lblStatus = new Label { Location = new Point(300, 20), AutoSize = true, ForeColor = Color.DimGray };

        // Help button
        var btnHelp = new Button
        {
            Text     = "❓ Help",
            Size     = new Size(80, 30),
            Location = new Point(bottom.Width - 460, 12),
            Anchor   = AnchorStyles.Right | AnchorStyles.Top,
            FlatStyle = FlatStyle.Flat,
        };
        btnHelp.Click += (_, _) => { using var f = new HelpForm(); f.ShowDialog(this); };

        // Save Excel Template button
        var btnExcelTpl = new Button
        {
            Text     = "Save Excel Template",
            Size     = new Size(150, 30),
            Location = new Point(bottom.Width - 370, 12),
            Anchor   = AnchorStyles.Right | AnchorStyles.Top,
            FlatStyle = FlatStyle.Flat,
        };
        btnExcelTpl.Click += SaveExcelTemplate;

        // Generate button
        btnGenerate = new Button
        {
            Text      = "Generate Presentation",
            Size      = new Size(172, 30),
            Location  = new Point(bottom.Width - 188, 12),
            Anchor    = AnchorStyles.Right | AnchorStyles.Top,
            Enabled   = false,
            BackColor = Color.FromArgb(0, 120, 215),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font      = new Font("Segoe UI", 9.5f, FontStyle.Bold),
        };
        btnGenerate.FlatAppearance.BorderSize = 0;
        btnGenerate.Click += Generate;

        bottom.Controls.AddRange([progressBar, lblStatus, btnHelp, btnExcelTpl, btnGenerate]);

        Controls.Add(dgvPreview);
        Controls.Add(lblEntryCount);
        Controls.Add(top);
        Controls.Add(bottom);
    }

    private DataGridViewTextBoxColumn AddCol(string prop, string header, int width)
    {
        var col = new DataGridViewTextBoxColumn
        {
            DataPropertyName = prop, HeaderText = header, Width = width,
            SortMode = DataGridViewColumnSortMode.NotSortable,
        };
        dgvPreview.Columns.Add(col);
        return col;
    }

    private static Label MkLabel(string text) =>
        new() { Text = text, AutoSize = true,
                Anchor = AnchorStyles.Left | AnchorStyles.Top,
                Margin = new Padding(0, 5, 8, 0) };

    // ── Template list ────────────────────────────────────────────────────

    private void RefreshTemplateList()
    {
        var configs  = ConfigManager.LoadAll();
        var previous = _activeConfig?.ConfigName;
        cmbTemplates.DataSource = null;
        cmbTemplates.DataSource = configs;

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

    private void CmbTemplates_Changed(object? sender, EventArgs e)
    {
        _activeConfig = cmbTemplates.SelectedItem as TemplateConfig;
        UpdateTemplateSummary();
        UpdateGenerateButton();
    }

    private void UpdateTemplateSummary()
    {
        if (_activeConfig == null) { lblTemplateSummary.Text = ""; return; }
        var c = _activeConfig;
        bool ok = File.Exists(c.TemplatePath);
        lblTemplateSummary.ForeColor = ok ? Color.DimGray : Color.OrangeRed;
        lblTemplateSummary.Text =
            $"{(ok ? "" : "⚠ File missing — ")}File: {Path.GetFileName(c.TemplatePath)}   " +
            $"Slides: {c.StaticSlides.Count} static / " +
            $"{(c.IndexSlide != null ? 1 : 0)} index / " +
            $"{(c.WordSlide  != null ? 1 : 0)} word   " +
            $"Words/page: {c.WordsPerIndexPage}   Hyperlinks: {(c.HyperlinkIndex ? "On" : "Off")}";
    }

    // ── Excel loading ────────────────────────────────────────────────────

    private void LoadExcel(object? sender, EventArgs e)
    {
        using var dlg = new OpenFileDialog
        {
            Title = "Select Word List Spreadsheet", Filter = "Excel Files|*.xlsx",
        };
        if (dlg.ShowDialog() != DialogResult.OK) return;

        _excelPath = dlg.FileName;
        txtExcelPath.Text = dlg.FileName;
        SetStatus("Reading spreadsheet…");

        try
        {
            _entries = ExcelReader.Read(_excelPath);
            dgvPreview.AutoGenerateColumns = false;
            dgvPreview.DataSource = null;
            dgvPreview.DataSource = _entries;

            lblEntryCount.Text =
                $"{_entries.Count} word(s) loaded — " +
                $"{_entries.Count(x => x.HasImage)} with images, " +
                $"{_entries.Count(x => x.HasAudio)} with audio";
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

    // ── Asset cell click ─────────────────────────────────────────────────

    private void Grid_CellClick(object? sender, DataGridViewCellEventArgs e)
    {
        if (e.RowIndex < 0 || e.RowIndex >= _entries.Count) return;
        bool isImg = e.ColumnIndex == colImage.Index;
        bool isAud = e.ColumnIndex == colAudio.Index;
        if (!isImg && !isAud) return;

        var entry = _entries[e.RowIndex];
        using var dlg = new OpenFileDialog();
        if (isImg) { dlg.Title = "Select Image"; dlg.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.gif;*.webp;*.bmp"; }
        else       { dlg.Title = "Select Audio"; dlg.Filter = "Audio Files|*.mp3;*.wav;*.m4a;*.ogg"; }

        if (dlg.ShowDialog() != DialogResult.OK) return;

        var asset = AssetData.FromPath(dlg.FileName, Path.GetDirectoryName(dlg.FileName)!);
        if (isImg) entry.Image = asset;
        else       entry.Audio = asset;

        dgvPreview.InvalidateRow(e.RowIndex);
    }

    private void Grid_CellFormatting(object? sender, DataGridViewCellFormattingEventArgs e)
    {
        if (e.RowIndex < 0 || e.RowIndex >= _entries.Count) return;
        var entry = _entries[e.RowIndex];

        if (e.ColumnIndex == colImage.Index)
        {
            e.Value = entry.ImageStatus;
            e.CellStyle.BackColor = entry.HasImage ? Color.FromArgb(220, 255, 220) : SystemColors.Window;
            e.FormattingApplied = true;
        }
        else if (e.ColumnIndex == colAudio.Index)
        {
            e.Value = entry.AudioStatus;
            e.CellStyle.BackColor = entry.HasAudio ? Color.FromArgb(210, 235, 255) : SystemColors.Window;
            e.FormattingApplied = true;
        }
    }

    // ── Save Excel Template ──────────────────────────────────────────────

    private void SaveExcelTemplate(object? sender, EventArgs e)
    {
        using var dlg = new SaveFileDialog
        {
            Title    = "Save Excel Template",
            Filter   = "Excel Files|*.xlsx",
            FileName = "WordList_Template.xlsx",
        };
        if (dlg.ShowDialog() != DialogResult.OK) return;

        try
        {
            ExcelTemplateWriter.Save(dlg.FileName);
            if (MessageBox.Show(
                    $"Template saved to:\n{dlg.FileName}\n\nOpen it now?",
                    "Saved", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    { FileName = dlg.FileName, UseShellExecute = true });
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to save template:\n\n{ex.Message}",
                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    // ── Generation ───────────────────────────────────────────────────────

    private async void Generate(object? sender, EventArgs e)
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

        var progress = new Progress<int>(n =>
        {
            progressBar.Value = Math.Min(n, progressBar.Maximum);
            lblStatus.Text    = $"Generating slide {n} of {_entries.Count}…";
        });

        try
        {
            var config  = _activeConfig;
            var entries = _entries.ToList();
            var output  = dlg.FileName;

            await Task.Run(() => PptxGenerator.Generate(config, output, entries, progress));

            SetStatus($"Done — {entries.Count} slide(s) generated.");

            if (MessageBox.Show(
                    $"Presentation created!\n\n{output}\n\nOpen now?",
                    "Success", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    { FileName = output, UseShellExecute = true });
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

    // ── Helpers ──────────────────────────────────────────────────────────

    private void UpdateGenerateButton() =>
        btnGenerate.Enabled = _activeConfig?.IsValid == true && _entries.Count > 0;

    private void SetGenerating(bool on)
    {
        btnGenerate.Enabled  = !on;
        btnLoadExcel.Enabled = !on;
        btnManage.Enabled    = !on;
        cmbTemplates.Enabled = !on;
        progressBar.Visible  = on;
        progressBar.Maximum  = Math.Max(1, _entries.Count);
        progressBar.Value    = 0;
    }

    private void SetStatus(string msg) => lblStatus.Text = msg;
}
