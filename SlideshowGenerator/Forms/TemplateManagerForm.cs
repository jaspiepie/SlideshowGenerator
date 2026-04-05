using LanguageCourseSlides.Models;
using LanguageCourseSlides.Services;

namespace LanguageCourseSlides.Forms;

public class TemplateManagerForm : Form
{
    private ListBox listTemplates = null!;
    private Label   lblDetail     = null!;
    private Button  btnNew        = null!;
    private Button  btnEdit       = null!;
    private Button  btnDelete     = null!;
    private Button  btnClose      = null!;

    private List<TemplateConfig> _configs = [];

    public TemplateManagerForm()
    {
        BuildUI();
        Reload();
    }

    private void BuildUI()
    {
        Text            = "Manage Templates";
        Size            = new Size(560, 440);
        MinimumSize     = new Size(460, 360);
        StartPosition   = FormStartPosition.CenterParent;
        Font            = new Font("Segoe UI", 9f);

        var split = new SplitContainer
        {
            Dock = DockStyle.Fill, Orientation = Orientation.Vertical,
            SplitterDistance = 200,
        };

        listTemplates = new ListBox { Dock = DockStyle.Fill, DisplayMember = nameof(TemplateConfig.ConfigName) };
        listTemplates.SelectedIndexChanged += (_, _) => UpdateDetail();
        listTemplates.DoubleClick          += (_, _) => Edit();
        split.Panel1.Controls.Add(listTemplates);

        var right = new TableLayoutPanel
        {
            Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, Padding = new Padding(6),
        };
        right.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        right.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        lblDetail = new Label
        {
            Dock = DockStyle.Fill, AutoSize = false, ForeColor = Color.DimGray,
            Font = new Font("Segoe UI", 8.5f),
        };
        right.Controls.Add(lblDetail, 0, 0);

        var btns = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, AutoSize = true,
        };
        btnNew    = Btn("New Template",    () => { using var f = new TemplateConfigForm();    if (f.ShowDialog() == DialogResult.OK) Reload(); });
        btnEdit   = Btn("Edit Selected",   Edit);
        btnDelete = Btn("Delete Selected", Delete);
        btnClose  = Btn("Close",           Close);

        btns.Controls.AddRange([btnNew, btnEdit, btnDelete, new Label { Height = 8 }, btnClose]);
        right.Controls.Add(btns, 0, 1);
        split.Panel2.Controls.Add(right);
        Controls.Add(split);
    }

    private static Button Btn(string text, Action onClick)
    {
        var b = new Button { Text = text, Width = 150, Height = 28, Margin = new Padding(0, 2, 0, 2) };
        b.Click += (_, _) => onClick();
        return b;
    }

    private void Reload()
    {
        _configs = ConfigManager.LoadAll();
        listTemplates.DataSource = null;
        listTemplates.DataSource = _configs;
        UpdateDetail();
    }

    private void UpdateDetail()
    {
        if (listTemplates.SelectedItem is not TemplateConfig c)
        {
            lblDetail.Text    = "Select a template to see details.";
            btnEdit.Enabled   = false;
            btnDelete.Enabled = false;
            return;
        }
        btnEdit.Enabled = btnDelete.Enabled = true;
        lblDetail.Text =
            $"File: {Path.GetFileName(c.TemplatePath)}\n" +
            $"Exists: {(File.Exists(c.TemplatePath) ? "Yes" : "⚠ Not found")}\n\n" +
            $"Slides: {c.StaticSlides.Count} static, " +
            $"{(c.IndexSlide != null ? 1 : 0)} index, " +
            $"{(c.WordSlide  != null ? 1 : 0)} word\n" +
            $"Words per index page: {c.WordsPerIndexPage}\n" +
            $"Index format: {c.IndexLineFormat}\n" +
            $"Hyperlinks: {(c.HyperlinkIndex ? "Yes" : "No")}\n\n" +
            $"Created: {c.CreatedAt:yyyy-MM-dd}";
    }

    private void Edit()
    {
        if (listTemplates.SelectedItem is not TemplateConfig sel) return;
        using var f = new TemplateConfigForm(sel);
        if (f.ShowDialog() == DialogResult.OK) Reload();
    }

    private void Delete()
    {
        if (listTemplates.SelectedItem is not TemplateConfig sel) return;
        if (MessageBox.Show($"Delete \"{sel.ConfigName}\"?\n\nThe .pptx file is not deleted.",
                "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;
        ConfigManager.Delete(sel);
        Reload();
    }
}
