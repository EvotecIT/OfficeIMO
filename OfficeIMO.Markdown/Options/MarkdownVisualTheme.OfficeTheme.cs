namespace OfficeIMO.Markdown;

public sealed partial class MarkdownVisualTheme {
    /// <summary>The immutable native theme retained by FromOfficeTheme, including typography, spacing, page and non-Markdown styles.</summary>
    public OfficeTheme? SourceTheme { get; set; }

    /// <summary>Applies Markdown palette/table edits without replacing native roles that were left unchanged.</summary>
    internal OfficeTheme? EffectiveSourceTheme {
        get {
            if (SourceTheme is not { } source) return null;
            MarkdownVisualTheme baseline = FromOfficeTheme(source);
            MarkdownVisualPalette original = baseline.PaletteSnapshot;
            MarkdownVisualPalette palette = PaletteSnapshot;
            MarkdownTableVisualStyle table = TableSnapshot;
            OfficeColor? ColorDelta(OfficeColor value, OfficeColor previous) {
                OfficeColor literal = value.Resolve(source);
                return literal == previous ? null : literal;
            }
            OfficeColor? accent = ColorDelta(palette.Accent, original.Accent);
            OfficeColor? heading = ColorDelta(palette.Heading, original.Heading);
            OfficeColor? text = ColorDelta(palette.Text, original.Text);
            OfficeColor? muted = ColorDelta(palette.MutedText, original.MutedText);
            OfficeColor? background = ColorDelta(palette.Background, original.Background);
            OfficeColor? surface = ColorDelta(palette.Surface, original.Surface);
            OfficeColor? border = ColorDelta(palette.Border, original.Border);
            OfficeColor? headerFill = ColorDelta(palette.TableHeaderBackground, original.TableHeaderBackground);
            OfficeColor? headerText = ColorDelta(palette.TableHeaderText, original.TableHeaderText);
            OfficeColor? bandFill = ColorDelta(palette.TableStripeBackground, original.TableStripeBackground);
            bool? emphasize = table.EmphasizeHeader != baseline.TableSnapshot.EmphasizeHeader ? table.EmphasizeHeader : null;
            bool? banding = table.UseRowStripes != baseline.TableSnapshot.UseRowStripes ? table.UseRowStripes : null;
            OfficeLength? borderWidth = table.BorderWidth != baseline.TableSnapshot.BorderWidth ? table.BorderWidth : null;
            OfficeLength? padding = table.CellPaddingY != baseline.TableSnapshot.CellPaddingY ? table.CellPaddingY : null;
            if (accent == null && heading == null && text == null && muted == null && background == null && surface == null &&
                border == null && headerFill == null && headerText == null && bandFill == null && emphasize == null &&
                banding == null && borderWidth == null && padding == null) return source;

            // HeaderBold=false in the native theme is independent of its fill/text. Only a later Markdown
            // emphasis edit opts into the Markdown convention of suppressing those colours together.
            if (emphasize == false) {
                headerFill = OfficeColor.Transparent;
                headerText = text ?? source.Palette.Text;
            }
            return source
                .WithPalette(source.Palette.With(accent1: accent, text: text, textMuted: muted, background: background))
                .WithTextStyle(source.TextStyle.With(headingColor: heading, linkColor: accent))
                .WithSurfaces(source.Surfaces.With(surface: surface, rule: border))
                .WithTableStyle(source.TableStyle.With(headerFill: headerFill, headerText: headerText, bandFill: bandFill,
                    borderColor: border, borderWidth: borderWidth, cellPadding: padding, banding: banding, headerBold: emphasize));
        }
    }

    /// <summary>Maps the shared palette and table appearance while retaining the complete native theme for format-specific rendering.</summary>
    public static MarkdownVisualTheme FromOfficeTheme(OfficeTheme theme) {
        ArgumentNullException.ThrowIfNull(theme);
        return new MarkdownVisualTheme {
            Name = theme.Name, HtmlStyle = MarkdownHtmlStyle.Clean, SourceTheme = theme,
            Palette = new MarkdownVisualPalette {
                Accent = theme.Palette.Accent1.Resolve(theme), Heading = theme.TextStyle.HeadingColor.Resolve(theme), Text = theme.Palette.Text.Resolve(theme),
                MutedText = theme.Palette.TextMuted.Resolve(theme), Background = theme.Palette.Background.Resolve(theme),
                Surface = theme.Surfaces.Surface.Resolve(theme), Border = theme.TableStyle.BorderColor.Resolve(theme),
                CodeBackground = theme.Surfaces.Surface.Resolve(theme), TableHeaderBackground = theme.TableStyle.HeaderFill.Resolve(theme),
                TableHeaderText = theme.TableStyle.HeaderText.Resolve(theme), TableStripeBackground = theme.TableStyle.BandFill.Resolve(theme)
            },
            Table = new MarkdownTableVisualStyle { UseRowStripes = theme.TableStyle.Banding, EmphasizeHeader = theme.TableStyle.HeaderBold,
                BorderWidth = theme.TableStyle.BorderWidth, CellPaddingY = theme.TableStyle.CellPadding,
                CellPaddingX = OfficeLength.Points(theme.TableStyle.CellPadding.ToPoints() * 1.4D) }
        };
    }
}
