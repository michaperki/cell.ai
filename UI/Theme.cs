using System.Drawing;
using System.Windows.Forms;

namespace SpreadsheetApp.UI;

/// <summary>
/// Centralized theme constants and button styling helpers.
/// </summary>
internal static class Theme
{
    // ── Accent ───────────────────────────────────────────────
    public static readonly Color Primary       = ColorTranslator.FromHtml("#2563EB");
    public static readonly Color PrimaryHover   = ColorTranslator.FromHtml("#1D4ED8");
    public static readonly Color Success        = ColorTranslator.FromHtml("#16A34A");
    public static readonly Color SuccessHover   = ColorTranslator.FromHtml("#15803D");
    public static readonly Color Danger         = ColorTranslator.FromHtml("#DC2626");
    public static readonly Color DangerHover    = ColorTranslator.FromHtml("#B91C1C");

    // ── Surfaces ─────────────────────────────────────────────
    public static readonly Color PanelBg        = Color.White;
    public static readonly Color PanelBorder    = Color.FromArgb(224, 224, 224);
    public static readonly Color InputBg        = Color.White;
    public static readonly Color SurfaceMuted   = Color.FromArgb(245, 245, 245);
    public static readonly Color LogBg          = Color.FromArgb(30, 30, 30);
    public static readonly Color LogFg          = Color.FromArgb(212, 212, 212);

    // ── Text ─────────────────────────────────────────────────
    public static readonly Color TextPrimary    = Color.FromArgb(30, 30, 30);
    public static readonly Color TextSecondary  = Color.FromArgb(107, 114, 128);
    public static readonly Color TextMuted      = Color.FromArgb(156, 163, 175);

    // ── Grid ─────────────────────────────────────────────────
    public static readonly Color GridLine       = Color.FromArgb(228, 228, 228);
    public static readonly Color HeaderBg       = Color.FromArgb(240, 240, 240);
    public static readonly Color HeaderFg       = Color.FromArgb(60, 60, 60);
    public static readonly Color SelectionBg    = Color.FromArgb(200, 220, 240);
    public static readonly Color FlashHighlight = Color.FromArgb(255, 255, 200);
    public static readonly Color AccentBlue     = Color.FromArgb(0, 120, 215);

    // ── Log colors (RichTextBox) ─────────────────────────────
    public static readonly Color LogCommand     = Color.FromArgb(229, 229, 229);
    public static readonly Color LogRationale   = Color.FromArgb(156, 163, 175);
    public static readonly Color LogError       = Color.FromArgb(248, 113, 113);
    public static readonly Color LogSuccess     = Color.FromArgb(74, 222, 128);
    public static readonly Color LogObservation = Color.FromArgb(103, 232, 249);
    public static readonly Color LogInfo        = Color.FromArgb(148, 163, 184);

    // ── Fonts ────────────────────────────────────────────────
    public static Font UI         => new("Segoe UI", 9F);
    public static Font UIBold     => new("Segoe UI", 9F, FontStyle.Bold);
    public static Font UISemiBold => new("Segoe UI Semibold", 9F);
    public static Font Mono       => new(MonoFamilyName, 9F);
    public static Font MonoSmall  => new(MonoFamilyName, 8.5F);

    private static string MonoFamilyName
    {
        get
        {
            // Prefer Cascadia Mono if available, fall back to Consolas
            using var test = new Font("Cascadia Mono", 9F);
            return test.Name.Equals("Cascadia Mono", StringComparison.OrdinalIgnoreCase)
                ? "Cascadia Mono" : "Consolas";
        }
    }

    // ── Button Helpers ───────────────────────────────────────

    public static void StylePrimary(Button btn)
    {
        btn.FlatStyle = FlatStyle.Flat;
        btn.BackColor = Primary;
        btn.ForeColor = Color.White;
        btn.FlatAppearance.BorderSize = 0;
        btn.FlatAppearance.MouseOverBackColor = PrimaryHover;
        btn.Font = UISemiBold;
        btn.Cursor = Cursors.Hand;
    }

    public static void StyleSuccess(Button btn)
    {
        btn.FlatStyle = FlatStyle.Flat;
        btn.BackColor = Success;
        btn.ForeColor = Color.White;
        btn.FlatAppearance.BorderSize = 0;
        btn.FlatAppearance.MouseOverBackColor = SuccessHover;
        btn.Font = UISemiBold;
        btn.Cursor = Cursors.Hand;
    }

    public static void StyleSecondary(Button btn)
    {
        btn.FlatStyle = FlatStyle.Flat;
        btn.BackColor = Color.Transparent;
        btn.ForeColor = Primary;
        btn.FlatAppearance.BorderColor = Primary;
        btn.FlatAppearance.BorderSize = 1;
        btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(30, Primary);
        btn.Font = UI;
        btn.Cursor = Cursors.Hand;
    }

    public static void StyleDanger(Button btn)
    {
        btn.FlatStyle = FlatStyle.Flat;
        btn.BackColor = Color.Transparent;
        btn.ForeColor = Danger;
        btn.FlatAppearance.BorderColor = Danger;
        btn.FlatAppearance.BorderSize = 1;
        btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(30, Danger);
        btn.Font = UI;
        btn.Cursor = Cursors.Hand;
    }

    public static void StyleGhost(Button btn)
    {
        btn.FlatStyle = FlatStyle.Flat;
        btn.BackColor = Color.Transparent;
        btn.ForeColor = TextSecondary;
        btn.FlatAppearance.BorderSize = 0;
        btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(20, 0, 0, 0);
        btn.Font = UI;
        btn.Cursor = Cursors.Hand;
    }

    /// <summary>
    /// Apply standard grid styling to any DataGridView (main grid or preview grids in dialogs).
    /// </summary>
    public static void StyleGrid(DataGridView dgv)
    {
        dgv.BorderStyle = BorderStyle.None;
        dgv.CellBorderStyle = DataGridViewCellBorderStyle.Single;
        dgv.GridColor = GridLine;
        dgv.BackgroundColor = Color.White;

        dgv.DefaultCellStyle.Font = UI;
        dgv.DefaultCellStyle.BackColor = Color.White;
        dgv.DefaultCellStyle.ForeColor = TextPrimary;
        dgv.DefaultCellStyle.SelectionBackColor = SelectionBg;
        dgv.DefaultCellStyle.SelectionForeColor = TextPrimary;
        dgv.DefaultCellStyle.Padding = new Padding(2, 0, 2, 0);

        dgv.EnableHeadersVisualStyles = false;
        dgv.ColumnHeadersDefaultCellStyle.BackColor = HeaderBg;
        dgv.ColumnHeadersDefaultCellStyle.ForeColor = HeaderFg;
        dgv.ColumnHeadersDefaultCellStyle.Font = UI;
        dgv.ColumnHeadersDefaultCellStyle.SelectionBackColor = HeaderBg;
        dgv.ColumnHeadersDefaultCellStyle.SelectionForeColor = HeaderFg;
        dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        dgv.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;

        dgv.RowHeadersDefaultCellStyle.BackColor = HeaderBg;
        dgv.RowHeadersDefaultCellStyle.ForeColor = HeaderFg;
        dgv.RowHeadersDefaultCellStyle.Font = UI;
        dgv.RowHeadersDefaultCellStyle.SelectionBackColor = HeaderBg;
        dgv.RowHeadersDefaultCellStyle.SelectionForeColor = HeaderFg;
        dgv.RowHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
    }
}

/// <summary>
/// Flat menu/toolbar renderer — removes gradients from menus and toolbars.
/// </summary>
internal class FlatToolStripRenderer : ToolStripProfessionalRenderer
{
    public FlatToolStripRenderer() : base(new FlatColorTable()) { }

    protected override void OnRenderMenuItemBackground(ToolStripItemRenderEventArgs e)
    {
        if (e.Item.Selected || e.Item.Pressed)
        {
            using var brush = new SolidBrush(Color.FromArgb(30, Theme.Primary));
            e.Graphics.FillRectangle(brush, new Rectangle(Point.Empty, e.Item.Size));
        }
    }

    private class FlatColorTable : ProfessionalColorTable
    {
        public override Color MenuStripGradientBegin => Theme.SurfaceMuted;
        public override Color MenuStripGradientEnd => Theme.SurfaceMuted;
        public override Color MenuItemSelected => Color.FromArgb(30, Theme.Primary);
        public override Color MenuItemSelectedGradientBegin => Color.FromArgb(30, Theme.Primary);
        public override Color MenuItemSelectedGradientEnd => Color.FromArgb(30, Theme.Primary);
        public override Color MenuItemPressedGradientBegin => Color.FromArgb(50, Theme.Primary);
        public override Color MenuItemPressedGradientEnd => Color.FromArgb(50, Theme.Primary);
        public override Color MenuBorder => Theme.PanelBorder;
        public override Color MenuItemBorder => Color.Transparent;
        public override Color ImageMarginGradientBegin => Color.White;
        public override Color ImageMarginGradientMiddle => Color.White;
        public override Color ImageMarginGradientEnd => Color.White;
        public override Color SeparatorDark => Theme.PanelBorder;
        public override Color SeparatorLight => Color.White;
        public override Color ToolStripDropDownBackground => Color.White;
        public override Color StatusStripGradientBegin => Theme.SurfaceMuted;
        public override Color StatusStripGradientEnd => Theme.SurfaceMuted;
    }
}
