namespace OfficeIMO.Markdown;

public sealed partial class MarkdownVisualTheme {
    /// <summary>The immutable native theme retained by FromOfficeTheme, including typography, spacing, page and non-Markdown styles.</summary>
    public OfficeTheme? SourceTheme { get; set; }

    /// <summary>Maps the shared palette and table appearance while retaining the complete native theme for format-specific rendering.</summary>
    public static MarkdownVisualTheme FromOfficeTheme(OfficeTheme theme) {
        ArgumentNullException.ThrowIfNull(theme);
        return new MarkdownVisualTheme {
            Name = theme.Name, HtmlStyle = MarkdownHtmlStyle.Clean, SourceTheme = theme,
            Palette = new MarkdownVisualPalette {
                Accent = theme.Palette.Accent1, Heading = theme.Palette.Text, Text = theme.Palette.Text,
                MutedText = theme.Palette.TextMuted, Background = theme.Palette.Background,
                Surface = theme.Surfaces.Surface, Border = theme.TableStyle.BorderColor.Resolve(theme),
                CodeBackground = theme.Surfaces.Surface, TableHeaderBackground = theme.TableStyle.HeaderFill.Resolve(theme),
                TableHeaderText = theme.TableStyle.HeaderText.Resolve(theme), TableStripeBackground = theme.TableStyle.BandFill.Resolve(theme)
            },
            Table = new MarkdownTableVisualStyle { UseRowStripes = theme.TableStyle.Banding, EmphasizeHeader = theme.TableStyle.HeaderBold,
                BorderWidth = theme.TableStyle.BorderWidth, CellPaddingY = theme.TableStyle.CellPadding,
                CellPaddingX = OfficeLength.Points(theme.TableStyle.CellPadding.ToPoints() * 1.4D) }
        };
    }
}
