using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using LanguageCourseSlides.Models;
using LanguageCourseSlides.Services;

namespace LanguageCourseSlides.Forms;

public class TemplateConfigForm : Form
{
    private TemplateConfig _config;
    public  TemplateConfig? ResultConfig { get; private set; }

    // Controls
    private TextBox      txtConfigName     = null!;
    private TextBox      txtTemplatePath   = null!;
    private Button       btnBrowse         = null!;
    private Label        lblSlideCount     = null!;
    private NumericUpDown numWordsPerPage  = null!;
    private TextBox      txtIndexFormat    = null!;
    private CheckBox     chkHyperlink      = null!;
    private DataGridView dgvSlides         = null!;
    private Button       btnSave           = null!;
    private Button       btnCancel         = null!;

    public TemplateConfigForm(TemplateConfig? existing = null)
    {
        _config = existing != null
            ? ConfigManager.Clone(existing)
            : new TemplateConfig();

        BuildUI();
        if (existing != null) Populate();
    }

    private void BuildUI()
    {
        Text            = "Template Configuration";
        Size            = new Size(740, 600);
        MinimumSize     = new Size(600, 500);
        StartPosition   = FormStartPosition.CenterParent;
        Font            = new System.Drawing.Font("Segoe UI", 9f);

        var outer = new TableLayoutPanel
        {
            Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3,
            Padding = new Padding(10),
        };
        outer.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        outer.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        outer.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        Controls.Add(outer);

        // ── Fields panel ────────────────────────────────────────────────
        var fields = new TableLayoutPanel
        {
            Dock = DockStyle.Fill, ColumnCount = 2, AutoSize = true,
        };
        fields.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150));
        fields.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

        int row = 0;

        // Name
        txtConfigName = new TextBox { Dock = DockStyle.Fill, PlaceholderText = "e.g. German A1 Standard" };
        AddFieldRow(fields, row++, "Template name:", txtConfigName);

        // File browser
        var browseRow = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
        txtTemplatePath = new TextBox { Width = 360, ReadOnly = true };
        btnBrowse = new Button { Text = "Browse…", AutoSize = true, Margin = new Padding(4, 0, 0, 0) };
        btnBrowse.Click += BrowseTemplate;
        browseRow.Controls.Add(txtTemplatePath);
        browseRow.Controls.Add(btnBrowse);
        AddFieldRow(fields, row++, "Template file (.pptx):", browseRow);

        lblSlideCount = new Label { AutoSize = true, ForeColor = Color.Gray };
        AddFieldRow(fields, row++, "", lblSlideCount);

        // Words per page
        numWordsPerPage = new NumericUpDown
        {
            Minimum = 1, Maximum = 200, Value = 20, Width = 70
        };
        AddFieldRow(fields, row++, "Words per index page:", numWordsPerPage);

        // Index format
        txtIndexFormat = new TextBox { Dock = DockStyle.Fill, Text = "{word} {type} {english}" };
        AddFieldRow(fields, row++, "Index line format:", txtIndexFormat);

        var hint = new Label
        {
            Text = "Each token becomes one table column (in the order listed).\n" +
                   "Available tokens:  {n}  {word}  {plural}  {type}  {english}\n" +
                   "Example: \"{n} {word} {plural} {type} {english}\" → 5 columns",
            AutoSize = true, ForeColor = Color.Gray,
            Font = new System.Drawing.Font(Font.FontFamily, 8),
        };
        AddFieldRow(fields, row++, "", hint);

        chkHyperlink = new CheckBox
        {
            Text = "Hyperlink index entries to their word slides", Checked = true, AutoSize = true
        };
        AddFieldRow(fields, row++, "", chkHyperlink);

        var shapeNote = new Label
        {
            Text =
                "Shape name conventions (name shapes in PowerPoint using the Selection Pane):\n" +
                "  shape_audio — audio trigger    shape_image — image placeholder\n" +
                "  shape_next  — next slide link  shape_index — back to index link\n" +
                "  {{Index}}   — text box on index slide where word list is injected",
            AutoSize = true, ForeColor = Color.DimGray,
            Font = new System.Drawing.Font(Font.FontFamily, 8.5f),
            Padding = new Padding(0, 6, 0, 0),
        };
        AddFieldRow(fields, row++, "", shapeNote);
        fields.SetColumnSpan(shapeNote, 2);

        outer.Controls.Add(fields, 0, 0);

        // ── Slide grid ──────────────────────────────────────────────────
        var grp = new GroupBox { Text = "Slide Roles", Dock = DockStyle.Fill };
        dgvSlides = BuildGrid();
        grp.Controls.Add(dgvSlides);
        outer.Controls.Add(grp, 0, 1);

        // ── Buttons ─────────────────────────────────────────────────────
        var btnRow = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft,
            AutoSize = true, Padding = new Padding(0, 6, 0, 0),
        };
        btnCancel = new Button { Text = "Cancel", Width = 80, DialogResult = DialogResult.Cancel };
        btnSave   = new Button { Text = "Save",   Width = 80 };
        btnSave.Click += Save;
        btnRow.Controls.Add(btnCancel);
        btnRow.Controls.Add(btnSave);
        outer.Controls.Add(btnRow, 0, 2);

        AcceptButton = btnSave;
        CancelButton = btnCancel;
    }

    private static void AddFieldRow(TableLayoutPanel panel, int row, string label, System.Windows.Forms.Control ctrl)
    {
        panel.Controls.Add(
            new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Right | AnchorStyles.Top,
                        TextAlign = ContentAlignment.MiddleRight, Margin = new Padding(0,5,6,0) },
            0, row);
        panel.Controls.Add(ctrl, 1, row);
    }

    private DataGridView BuildGrid()
    {
        var dgv = new DataGridView
        {
            Dock = DockStyle.Fill, AllowUserToAddRows = false, AllowUserToDeleteRows = false,
            RowHeadersVisible = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            BackgroundColor = SystemColors.Control, BorderStyle = BorderStyle.None,
        };

        dgv.Columns.Add(new DataGridViewTextBoxColumn
            { Name = "colNum", HeaderText = "#", Width = 36, ReadOnly = true });
        dgv.Columns.Add(new DataGridViewTextBoxColumn
            { Name = "colLabel", HeaderText = "Label (your reference)", Width = 200 });

        var roleCol = new DataGridViewComboBoxColumn
            { Name = "colRole", HeaderText = "Role", Width = 80 };
        roleCol.Items.AddRange("Static", "Index", "Word");
        dgv.Columns.Add(roleCol);

        dgv.EditingControlShowing += (s, e) =>
        {
            if (dgv.CurrentCell?.OwningColumn.Name == "colRole" && e.Control is ComboBox cb)
                cb.DropDownStyle = ComboBoxStyle.DropDownList;
        };

        return dgv;
    }

    // ── Events ────────────────────────────────────────────────────────────

    private void BrowseTemplate(object? sender, EventArgs e)
    {
        using var dlg = new OpenFileDialog
        {
            Filter = "PowerPoint Files|*.pptx", Title = "Select Template File"
        };
        if (dlg.ShowDialog() != DialogResult.OK) return;
        _config.TemplatePath = dlg.FileName;
        txtTemplatePath.Text = dlg.FileName;
        LoadSlides(dlg.FileName);
    }

    private void LoadSlides(string path)
    {
        dgvSlides.Rows.Clear();
        try
        {
            using var prs  = PresentationDocument.Open(path, isEditable: false);
            int count = prs.PresentationPart!.SlideParts.Count();
            for (int i = 0; i < count; i++)
            {
                var ex  = _config.Slides.FirstOrDefault(s => s.SlideIndex == i);
                int idx = dgvSlides.Rows.Add();
                var r   = dgvSlides.Rows[idx];
                r.Cells["colNum"].Value   = i + 1;
                r.Cells["colLabel"].Value = ex?.Label ?? $"Slide {i + 1}";
                r.Cells["colRole"].Value  = ex?.Role.ToString() ?? "Static";
            }
            lblSlideCount.Text      = $"{count} slide(s) found.";
            lblSlideCount.ForeColor = Color.DarkGreen;
        }
        catch (Exception ex)
        {
            lblSlideCount.Text      = $"Error: {ex.Message}";
            lblSlideCount.ForeColor = Color.Red;
        }
    }

    private void Save(object? sender, EventArgs e)
    {
        if (string.IsNullOrWhiteSpace(txtConfigName.Text))
        { Warn("Please enter a template name."); return; }

        if (!File.Exists(_config.TemplatePath))
        { Warn("Please select a valid .pptx template file."); return; }

        var slides = CollectSlides();
        if (slides.Count(s => s.Role == SlideRole.Word) != 1)
        { Warn("Exactly one slide must be assigned the 'Word' role."); return; }
        if (slides.Count(s => s.Role == SlideRole.Index) > 1)
        { Warn("At most one slide may be the 'Index' role (it is cloned if overflow occurs)."); return; }

        _config.ConfigName       = txtConfigName.Text.Trim();
        _config.Slides           = slides;
        _config.WordsPerIndexPage = (int)numWordsPerPage.Value;
        _config.IndexLineFormat  = txtIndexFormat.Text.Trim().IfEmpty("{word}");
        _config.HyperlinkIndex   = chkHyperlink.Checked;

        ConfigManager.Save(_config);
        ResultConfig = _config;
        DialogResult = DialogResult.OK;
        Close();
    }

    // ── Helpers ───────────────────────────────────────────────────────────

    private void Populate()
    {
        txtConfigName.Text      = _config.ConfigName;
        txtTemplatePath.Text    = _config.TemplatePath;
        numWordsPerPage.Value   = Math.Max(1, _config.WordsPerIndexPage);
        txtIndexFormat.Text     = _config.IndexLineFormat;
        chkHyperlink.Checked    = _config.HyperlinkIndex;
        if (File.Exists(_config.TemplatePath))
            LoadSlides(_config.TemplatePath);
    }

    private List<SlideDefinition> CollectSlides()
    {
        var list = new List<SlideDefinition>();
        foreach (DataGridViewRow r in dgvSlides.Rows)
        {
            if (r.IsNewRow) continue;
            Enum.TryParse<SlideRole>(r.Cells["colRole"].Value?.ToString(), out var role);
            list.Add(new SlideDefinition
            {
                SlideIndex = (int)r.Cells["colNum"].Value! - 1,
                Label      = r.Cells["colLabel"].Value?.ToString() ?? "",
                Role       = role,
            });
        }
        return list.OrderBy(d => d.SlideIndex).ToList();
    }

    private static void Warn(string msg) =>
        MessageBox.Show(msg, "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
}

internal static class StringExt
{
    public static string IfEmpty(this string s, string fallback) =>
        string.IsNullOrWhiteSpace(s) ? fallback : s;
}
