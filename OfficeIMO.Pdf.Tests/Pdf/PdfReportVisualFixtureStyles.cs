using OfficeIMO.Pdf;

namespace OfficeIMO.Tests.Pdf;

/// <summary>Explicit typography for the fixed report snapshots and two-page statement fixture.</summary>
internal static class PdfReportVisualFixtureStyles {
    /// <summary>
    /// Keeps the report fixtures' measured rhythm independent of the default Office theme.
    /// Page frames, body sizes, table styles and block colors remain supplied by each fixture.
    /// </summary>
    internal static PdfCreateOptions Apply(PdfCreateOptions options) {
        double leading = options.DefaultFontSize * 1.4D;
        options.DefaultTextColor ??= OfficeColor.Black;
        options.FooterTextColor = OfficeColor.Black;
        return options.ApplyTheme(new PdfTheme {
            ParagraphStyle = new PdfParagraphStyle {
                LineHeight = 1.4D,
                SpacingBefore = 0D,
                SpacingAfter = leading * 0.3D
            },
            HeadingStyles = new PdfHeadingStyles {
                Level1 = Heading(24D, options),
                Level3 = Heading(14D, options)
            },
            ListStyle = new PdfListStyle {
                FontSize = options.DefaultFontSize,
                LineHeight = 1.4D,
                LeftIndent = 0D,
                SpacingBefore = 0D,
                ItemSpacing = leading * 0.15D,
                SpacingAfter = leading * 0.15D,
                Color = options.DefaultTextColor,
                MarkerColor = options.DefaultTextColor
            }
        });
    }

    private static PdfHeadingStyle Heading(double size, PdfCreateOptions options) => new PdfHeadingStyle {
        Font = options.DefaultFont,
        FontSize = size,
        Bold = true,
        Color = options.DefaultTextColor,
        LineHeight = 1.25D,
        SpacingBefore = 0D,
        SpacingAfter = size * 1.25D * 0.25D,
        KeepWithNext = true
    };
}
