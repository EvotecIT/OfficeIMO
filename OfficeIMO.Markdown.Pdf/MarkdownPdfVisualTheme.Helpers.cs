using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Markdown.Pdf;

public sealed partial class MarkdownPdfVisualTheme {
    private static PdfCore.PdfTableStyle TechnicalTable() {
        PdfCore.PdfTableStyle style = MarkdownTable(PdfCore.TableStyles.Light(), 9.75, 9.75, 1.2, 6, 5, 6, 10);
        style.HeaderFill = PdfCore.PdfColor.FromRgb(15, 23, 42);
        style.HeaderTextColor = PdfCore.PdfColor.White;
        style.RowStripeFill = PdfCore.PdfColor.FromRgb(248, 250, 252);
        style.BorderColor = PdfCore.PdfColor.FromRgb(203, 213, 225);
        style.RowSeparatorColor = PdfCore.PdfColor.FromRgb(226, 232, 240);
        style.BorderWidth = 0.45;
        style.HeaderSeparatorWidth = 0.75;
        return style;
    }

    private static PdfCore.PdfTableStyle GitHubTable() {
        PdfCore.PdfTableStyle style = MarkdownTable(PdfCore.TableStyles.Light(), 10, 10, 1.22, 6, 5, 6, 10);
        style.HeaderFill = PdfCore.PdfColor.FromRgb(246, 248, 250);
        style.HeaderTextColor = PdfCore.PdfColor.FromRgb(36, 41, 47);
        style.TextColor = PdfCore.PdfColor.FromRgb(36, 41, 47);
        style.RowStripeFill = null;
        style.BorderColor = PdfCore.PdfColor.FromRgb(208, 215, 222);
        style.BorderWidth = 0.5;
        return style;
    }

    private static PdfCore.PdfTableStyle CompactTable() {
        PdfCore.PdfTableStyle style = MarkdownTable(PdfCore.TableStyles.ListTable1Light(), 9, 9, 1.12, 4, 3, 4, 7);
        style.RowSeparatorColor = PdfCore.PdfColor.FromRgb(226, 232, 240);
        return style;
    }

    private static PdfCore.PdfTableStyle ReportTable() {
        PdfCore.PdfTableStyle style = MarkdownTable(PdfCore.TableStyles.Light(), 9.25, 9.25, 1.18, 5, 4, 6, 10);
        style.HeaderFill = PdfCore.PdfColor.FromRgb(30, 64, 175);
        style.HeaderTextColor = PdfCore.PdfColor.White;
        style.RowStripeFill = PdfCore.PdfColor.FromRgb(239, 246, 255);
        style.BorderColor = PdfCore.PdfColor.FromRgb(191, 219, 254);
        style.BorderWidth = 0.45;
        return style;
    }

    private static PdfCore.PdfTableStyle FrontMatterTable() {
        PdfCore.PdfTableStyle style = MarkdownTable(PdfCore.TableStyles.Minimal(), 10, 10, 1.18, 6, 5, 0, 10);
        style.HeaderFill = PdfCore.PdfColor.FromRgb(241, 245, 249);
        style.HeaderTextColor = PdfCore.PdfColor.FromRgb(15, 23, 42);
        style.RowStripeFill = PdfCore.PdfColor.FromRgb(248, 250, 252);
        style.BorderColor = PdfCore.PdfColor.FromRgb(226, 232, 240);
        return style;
    }

    private static PdfCore.PdfTableStyle MarkdownTable(PdfCore.PdfTableStyle style, double fontSize, double headerFontSize, double lineHeight, double paddingX, double paddingY, double spacingBefore, double spacingAfter) {
        style.FontSize = fontSize;
        style.HeaderFontSize = headerFontSize;
        style.LineHeight = lineHeight;
        style.CellPaddingX = paddingX;
        style.CellPaddingY = paddingY;
        style.SpacingBefore = spacingBefore;
        style.SpacingAfter = spacingAfter;
        style.AutoFitColumns = true;
        return style;
    }

    private static PdfCore.PdfTableStyle DefinitionTable(int borderR, int borderG, int borderB, double borderWidth, double paddingX, double paddingY) {
        PdfCore.PdfTableStyle style = PdfCore.TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.CellPaddingX = paddingX;
        style.CellPaddingY = paddingY;
        style.BorderColor = Color(borderR, borderG, borderB);
        style.BorderWidth = borderWidth;
        style.SpacingBefore = 3;
        style.SpacingAfter = 8;
        return style;
    }

    private static PdfCore.PdfTableStyle ChecklistTable(int borderR, int borderG, int borderB, int textR, int textG, int textB, double fontSize, double paddingY) {
        PdfCore.PdfTableStyle style = PdfCore.TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.RepeatHeaderRowCount = 0;
        style.BorderColor = null;
        style.BorderWidth = 0;
        style.RowStripeFill = null;
        style.RowSeparatorColor = Color(borderR, borderG, borderB);
        style.RowSeparatorWidth = 0.35;
        style.TextColor = Color(textR, textG, textB);
        style.FontSize = fontSize;
        style.CellPaddingX = 3;
        style.CellPaddingY = paddingY;
        style.ColumnWidthPoints = new System.Collections.Generic.List<double?> { 20D, null };
        style.Alignments = new System.Collections.Generic.List<PdfCore.PdfColumnAlign> { PdfCore.PdfColumnAlign.Center, PdfCore.PdfColumnAlign.Left };
        style.VerticalAlignments = new System.Collections.Generic.List<PdfCore.PdfCellVerticalAlign> { PdfCore.PdfCellVerticalAlign.Top, PdfCore.PdfCellVerticalAlign.Top };
        style.BodyColumnFills = new System.Collections.Generic.List<PdfCore.PdfColor?> { null, null };
        style.SpacingBefore = 3;
        style.SpacingAfter = 8;
        return style;
    }

    private static PdfCore.PanelStyle Panel(int backgroundR, int backgroundG, int backgroundB, int borderR, int borderG, int borderB, double borderWidth, double paddingX, double paddingY, double spacingBefore, double spacingAfter) => new PdfCore.PanelStyle {
        Background = Color(backgroundR, backgroundG, backgroundB),
        BorderColor = Color(borderR, borderG, borderB),
        BorderWidth = borderWidth,
        PaddingX = paddingX,
        PaddingY = paddingY,
        SpacingBefore = spacingBefore,
        SpacingAfter = spacingAfter,
        KeepTogether = true
    };

    private static PdfCore.PanelStyle PanelWithLeftRule(int backgroundR, int backgroundG, int backgroundB, int borderR, int borderG, int borderB, int leftR, int leftG, int leftB, double borderWidth, double leftWidth, double paddingX, double paddingY, double spacingBefore, double spacingAfter) {
        PdfCore.PanelStyle style = Panel(backgroundR, backgroundG, backgroundB, borderR, borderG, borderB, borderWidth, paddingX, paddingY, spacingBefore, spacingAfter);
        style.LeftBorder = new PdfCore.PdfPanelBorder {
            Color = Color(leftR, leftG, leftB),
            Width = leftWidth
        };
        return style;
    }

    private static PdfCore.PdfColor GetCalloutColor(string? kind) {
        string normalized = (kind ?? string.Empty).Trim().ToLowerInvariant();
        switch (normalized) {
            case "warning":
            case "warn":
            case "caution":
            case "attention":
                return PdfCore.PdfColor.FromRgb(217, 119, 6);
            case "danger":
            case "error":
            case "failure":
                return PdfCore.PdfColor.FromRgb(220, 38, 38);
            case "success":
            case "check":
            case "done":
                return PdfCore.PdfColor.FromRgb(22, 163, 74);
            case "tip":
            case "hint":
                return PdfCore.PdfColor.FromRgb(14, 165, 233);
            case "important":
                return PdfCore.PdfColor.FromRgb(124, 58, 237);
            default:
                return PdfCore.PdfColor.FromRgb(37, 99, 235);
        }
    }

    private static PdfCore.PdfColor Color(int r, int g, int b) => PdfCore.PdfColor.FromRgb((byte)r, (byte)g, (byte)b);

    private static double? ValidateOptionalPositive(double? value, string propertyName) {
        if (value.HasValue && (double.IsNaN(value.Value) || double.IsInfinity(value.Value) || value.Value <= 0)) {
            throw new ArgumentOutOfRangeException(propertyName, "Markdown PDF visual theme sizes must be positive finite values.");
        }

        return value;
    }

    private static string NormalizeThemeName(string? value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return string.Empty;
        }

        var builder = new System.Text.StringBuilder(value!.Length);
        for (int i = 0; i < value.Length; i++) {
            char ch = value[i];
            if (char.IsLetterOrDigit(ch)) {
                builder.Append(char.ToLowerInvariant(ch));
            }
        }

        return builder.ToString();
    }
}
