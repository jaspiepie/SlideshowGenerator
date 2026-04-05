using SlideshowGenerator.Models;
using SlideshowGenerator.Services;

namespace SlideshowGenerator.Forms;

/// <summary>
/// Lists all saved TemplateConfigs and allows New / Edit / Delete.
/// </summary>
public class TemplateManagerForm : Form
{
    private ListBox   lstTemplates = null!;
    private Button    btnNew       = null!;
    private Button    btnEdit      = null!;
    private Button    btnDelete    = null!;
    private Button    btnClose     = null!;
    private Label     lblDetails   = null!;

    private List<TemplateConfig> _configs = [];

    public TemplateManagerForm()
    {
        BuildUI();
        Reload();
    }

    // ── UI ───────────────────────────────────────────────────────────────

    private void BuildUI()
    {
        Text            = "Manage Templates";
        Size            = new Size(520, 400);
        MinimumSize     = new Size(420, 320);
        StartPosition   = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.Sizable;
        Font            = new Font("Segoe UI", 9f);

        var split = new SplitContainer
        {
            Dock        = DockStyle.Fill,
            Orientation = Orientation.Horizontal,
            SplitterDistance = 260,
            Panel2MinSize    = 80,
        };

        lstTemplates = new ListBox
        {
            Dock          = DockStyle.Fill,
            DisplayMember = nameof(TemplateConfig.ConfigName),
            BorderStyle   = BorderStyle.None,
            ItemHeight    = 22,
        };
        lstTemplates.SelectedIndexChanged += LstTemplates_SelectedIndexChanged;
        lstTemplates.DoubleClick          += (s, e) => EditSelected();

        split.Panel1.Controls.Add(lstTemplates);

        lblDetails = new Label
        {
            Dock      = DockStyle.Fill,
            Padding   = new Padding(10, 8, 10, 4),
            ForeColor = Color.DimGray,
        };
        split.Panel2.Controls.Add(lblDetails);

        var btnPanel = new FlowLayoutPanel
        {
            Dock          = DockStyle.Bottom,
            Height        = 44,
            FlowDirection = FlowDirection.RightToLeft,
            Padding       = new Padding(8, 6, 8, 4),
        };

        btnClose  = new Button { Text = "Close",  Width = 80, DialogResult = DialogResult.Cancel };
        btnDelete = new Button { Text = "Delete",  Width = 80, Enabled = false };
        btnEdit   = new Button { Text = "Edit…",   Width = 80, Enabled = false };
        btnNew    = new Button { Text = "New…",    Width = 80 };

        btnNew.Click    += (s, e) => CreateNew();
        btnEdit.Click   += (s, e) => EditSelected();
        btnDelete.Click += (s, e) => DeleteSelected();

        btnPanel.Controls.AddRange([btnClose, btnDelete, btnEdit, btnNew]);

        Controls.Add(split);
        Controls.Add(btnPanel);
    }

    // ── actions ──────────────────────────────────────────────────────────

    private void CreateNew()
    {
        using var form = new TemplateConfigForm();
        if (form.ShowDialog(this) == DialogResult.OK) Reload();
    }

    private void EditSelected()
    {
        if (lstTemplates.SelectedItem is not TemplateConfig selected) return;
        using var form = new TemplateConfigForm(selected);
        if (form.ShowDialog(this) == DialogResult.OK) Reload();
    }

    private void DeleteSelected()
    {
        if (lstTemplates.SelectedItem is not TemplateConfig selected) return;
        if (MessageBox.Show(
                $"Delete template '{selected.ConfigName}'?",
                "Confirm Delete",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning) != DialogResult.Yes) return;

        ConfigManager.Delete(selected);
        Reload();
    }

    // ── helpers ──────────────────────────────────────────────────────────

    private void Reload()
    {
        _configs = ConfigManager.LoadAll();
        lstTemplates.DataSource = null;
        lstTemplates.DataSource = _configs;
    }

    private void LstTemplates_SelectedIndexChanged(object? sender, EventArgs e)
    {
        var sel = lstTemplates.SelectedItem as TemplateConfig;
        btnEdit.Enabled   = sel != null;
        btnDelete.Enabled = sel != null;

        if (sel == null) { lblDetails.Text = ""; return; }

        var staticCount = sel.StaticSlides.Count;
        var hasIndex    = sel.IndexSlide != null;
        var hasWord     = sel.WordSlide  != null;

        lblDetails.Text =
            $"File:     {sel.TemplatePath}\n" +
            $"Slides:   {sel.Slides.Count} total — " +
                $"{staticCount} static, " +
                $"{(hasIndex ? 1 : 0)} index, " +
                $"{(hasWord  ? 1 : 0)} word\n" +
            $"Index format:  {sel.IndexLineFormat}\n" +
            $"Hyperlinks:    {(sel.HyperlinkIndex ? "Yes" : "No")}\n" +
            $"Image placeholder: {sel.ImagePlaceholder}\n" +
            $"Audio placeholder: {sel.AudioPlaceholder}";
    }
}
