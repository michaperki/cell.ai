using System.Drawing;
using System.Windows.Forms;

namespace SpreadsheetApp.UI;

/// <summary>
/// Centralized theme with swappable light/dark palettes and button styling helpers.
/// </summary>
internal static class Theme
{
    private static ThemePalette _palette = ThemePalette.Light;
    public static bool IsDark => _palette == ThemePalette.Dark;

    public static void SetLight() => _palette = ThemePalette.Light;
    public static void SetDark() => _palette = ThemePalette.Dark;
    public static void Toggle() { if (IsDark) SetLight(); else SetDark(); }

    // ── Accent (shared across themes) ───────────────────────
    public static Color Primary       => ColorTranslator.FromHtml("#2563EB");
    public static Color PrimaryHover  => ColorTranslator.FromHtml("#1D4ED8");
    public static Color Success       => ColorTranslator.FromHtml("#16A34A");
    public static Color SuccessHover  => ColorTranslator.FromHtml("#15803D");
    public static Color Danger        => ColorTranslator.FromHtml("#DC2626");
    public static Color DangerHover   => ColorTranslator.FromHtml("#B91C1C");

    // ── Surfaces ─────────────────────────────────────────────
    public static Color PanelBg       => _palette.PanelBg;
    public static Color PanelBorder   => _palette.PanelBorder;
    public static Color InputBg       => _palette.InputBg;
    public static Color SurfaceMuted  => _palette.SurfaceMuted;
    public static Color LogBg         => _palette.LogBg;
    public static Color LogFg         => _palette.LogFg;
    public static Color FormBg        => _palette.FormBg;

    // ── Text ─────────────────────────────────────────────────
    public static Color TextPrimary   => _palette.TextPrimary;
    public static Color TextSecondary => _palette.TextSecondary;
    public static Color TextMuted     => _palette.TextMuted;

    // ── Grid ─────────────────────────────────────────────────
    public static Color GridLine       => _palette.GridLine;
    public static Color HeaderBg       => _palette.HeaderBg;
    public static Color HeaderFg       => _palette.HeaderFg;
    public static Color SelectionBg    => _palette.SelectionBg;
    public static Color FlashHighlight => _palette.FlashHighlight;
    public static Color AccentBlue     => Color.FromArgb(0, 120, 215);

    // ── Log colors (RichTextBox) ─────────────────────────────
    public static Color LogCommand     => _palette.LogCommand;
    public static Color LogRationale   => _palette.LogRationale;
    public static Color LogError       => Color.FromArgb(248, 113, 113);
    public static Color LogSuccess     => Color.FromArgb(74, 222, 128);
    public static Color LogObservation => Color.FromArgb(103, 232, 249);
    public static Color LogInfo        => _palette.LogInfo;

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

    public static Panel WrapWithFocusBorder(TextBox tb)
    {
        var wrapper = new Panel
        {
            Dock = tb.Dock,
            Height = tb.Height + 2,
            Width = tb.Width + 2,
            Padding = new Padding(1),
            BackColor = PanelBorder,
        };
        tb.Dock = DockStyle.Fill;
        tb.BorderStyle = BorderStyle.None;
        wrapper.Controls.Add(tb);
        tb.Enter += (_, __) => wrapper.BackColor = Primary;
        tb.Leave += (_, __) => wrapper.BackColor = PanelBorder;
        return wrapper;
    }

    public static void StyleGrid(DataGridView dgv)
    {
        dgv.BorderStyle = BorderStyle.None;
        dgv.CellBorderStyle = DataGridViewCellBorderStyle.Single;
        dgv.GridColor = GridLine;
        dgv.BackgroundColor = PanelBg;

        dgv.DefaultCellStyle.Font = UI;
        dgv.DefaultCellStyle.BackColor = PanelBg;
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

    /// <summary>
    /// Re-apply theme colors to an entire form and its controls after a theme switch.
    /// </summary>
    public static void ApplyToForm(Form form)
    {
        form.BackColor = FormBg;
        ApplyToControls(form.Controls);
        form.Invalidate(true);
        form.Refresh();
    }

    private static void ApplyToControls(Control.ControlCollection controls)
    {
        foreach (Control c in controls)
        {
            switch (c)
            {
                case DataGridView dgv:
                    StyleGrid(dgv);
                    break;
                case MenuStrip ms:
                    ms.BackColor = SurfaceMuted;
                    ms.ForeColor = TextPrimary;
                    ms.Renderer = new FlatToolStripRenderer();
                    break;
                case StatusStrip ss:
                    ss.BackColor = SurfaceMuted;
                    ss.Renderer = new FlatToolStripRenderer();
                    foreach (ToolStripItem item in ss.Items)
                        item.ForeColor = TextSecondary;
                    break;
                case RichTextBox rtb when rtb.ReadOnly:
                    rtb.BackColor = LogBg;
                    rtb.ForeColor = LogFg;
                    break;
                case TabControl tc:
                    tc.Invalidate();
                    break;
            }

            // Recurse
            if (c.Controls.Count > 0)
                ApplyToControls(c.Controls);
        }
    }
}

// ── Theme Palette ────────────────────────────────────────────

internal class ThemePalette
{
    // Surfaces
    public Color PanelBg, PanelBorder, InputBg, SurfaceMuted, LogBg, LogFg, FormBg;
    // Text
    public Color TextPrimary, TextSecondary, TextMuted;
    // Grid
    public Color GridLine, HeaderBg, HeaderFg, SelectionBg, FlashHighlight;
    // Log
    public Color LogCommand, LogRationale, LogInfo;

    public static readonly ThemePalette Light = new()
    {
        PanelBg       = Color.White,
        PanelBorder   = Color.FromArgb(224, 224, 224),
        InputBg       = Color.White,
        SurfaceMuted  = Color.FromArgb(245, 245, 245),
        LogBg         = Color.FromArgb(30, 30, 30),
        LogFg         = Color.FromArgb(212, 212, 212),
        FormBg        = Color.White,
        TextPrimary   = Color.FromArgb(30, 30, 30),
        TextSecondary = Color.FromArgb(107, 114, 128),
        TextMuted     = Color.FromArgb(156, 163, 175),
        GridLine      = Color.FromArgb(228, 228, 228),
        HeaderBg      = Color.FromArgb(240, 240, 240),
        HeaderFg      = Color.FromArgb(60, 60, 60),
        SelectionBg   = Color.FromArgb(200, 220, 240),
        FlashHighlight = Color.FromArgb(255, 255, 200),
        LogCommand    = Color.FromArgb(229, 229, 229),
        LogRationale  = Color.FromArgb(156, 163, 175),
        LogInfo       = Color.FromArgb(148, 163, 184),
    };

    public static readonly ThemePalette Dark = new()
    {
        PanelBg       = Color.FromArgb(30, 30, 30),
        PanelBorder   = Color.FromArgb(55, 55, 55),
        InputBg       = Color.FromArgb(42, 42, 42),
        SurfaceMuted  = Color.FromArgb(37, 37, 37),
        LogBg         = Color.FromArgb(22, 22, 22),
        LogFg         = Color.FromArgb(200, 200, 200),
        FormBg        = Color.FromArgb(30, 30, 30),
        TextPrimary   = Color.FromArgb(220, 220, 220),
        TextSecondary = Color.FromArgb(160, 160, 160),
        TextMuted     = Color.FromArgb(120, 120, 120),
        GridLine      = Color.FromArgb(55, 55, 55),
        HeaderBg      = Color.FromArgb(42, 42, 42),
        HeaderFg      = Color.FromArgb(180, 180, 180),
        SelectionBg   = Color.FromArgb(37, 55, 85),
        FlashHighlight = Color.FromArgb(80, 70, 20),
        LogCommand    = Color.FromArgb(200, 200, 200),
        LogRationale  = Color.FromArgb(130, 130, 130),
        LogInfo       = Color.FromArgb(140, 150, 160),
    };
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
        public override Color ImageMarginGradientBegin => Theme.PanelBg;
        public override Color ImageMarginGradientMiddle => Theme.PanelBg;
        public override Color ImageMarginGradientEnd => Theme.PanelBg;
        public override Color SeparatorDark => Theme.PanelBorder;
        public override Color SeparatorLight => Theme.PanelBg;
        public override Color ToolStripDropDownBackground => Theme.PanelBg;
        public override Color StatusStripGradientBegin => Theme.SurfaceMuted;
        public override Color StatusStripGradientEnd => Theme.SurfaceMuted;
    }
}
