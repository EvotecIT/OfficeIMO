using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Style = DocumentFormat.OpenXml.Wordprocessing.Style;

namespace OfficeIMO.Word {
    /// <summary>
    /// Builds the style definitions a new document receives from the default <see cref="OfficeTheme"/>.
    /// Loaded documents keep their own styles; nothing here is applied to them.
    /// </summary>
    internal static class WordThemeDefaults {
        internal const string TableStyleId = "OfficeIMOModern";

        // Horizontal padding gives text a little more room beside a cell rule.
        private const double TableHorizontalPaddingRatio = 1.4D;

        internal static OfficeTheme Theme => OfficeTheme.Default;

        private static string Hex(OfficeColor color) => color.Resolve(Theme).ToRgbHex();

        private static string HalfPoints(OfficeLength size) =>
            Math.Max(2, (int)Math.Round(size.ToPoints() * 2D, MidpointRounding.AwayFromZero)).ToString(CultureInfo.InvariantCulture);

        private static string Twips(double points) =>
            ((int)Math.Round(points * 20D, MidpointRounding.AwayFromZero)).ToString(CultureInfo.InvariantCulture);

        // Border sizes are in eighths of a point, between 2 and 96.
        private static uint EighthPoints(double points) =>
            (uint)Math.Min(96, Math.Max(2, Math.Round(points * 8D, MidpointRounding.AwayFromZero)));

        private static OfficeLength HeadingSize(int level) {
            OfficeThemeTypeScale scale = Theme.TypeScale;
            return level switch {
                1 => scale.Heading1,
                2 => scale.Heading2,
                3 => scale.Heading3,
                4 => scale.Heading4,
                5 => scale.Heading5,
                _ => scale.Heading6
            };
        }

        // One type family per role. Word themes carry the heading font as the major
        // font; when the theme uses one family for headings and body, headings use the
        // minor (body) font so a light display cut such as Calibri Light never stands in.
        private static RunFonts HeadingFonts() {
            bool sharedFamily = string.Equals(Theme.Fonts.Heading.Family, Theme.Fonts.Body.Family, StringComparison.OrdinalIgnoreCase);
            return sharedFamily
                ? new RunFonts {
                    AsciiTheme = ThemeFontValues.MinorHighAnsi,
                    HighAnsiTheme = ThemeFontValues.MinorHighAnsi,
                    EastAsiaTheme = ThemeFontValues.MinorEastAsia,
                    ComplexScriptTheme = ThemeFontValues.MinorBidi
                }
                : new RunFonts {
                    AsciiTheme = ThemeFontValues.MajorHighAnsi,
                    HighAnsiTheme = ThemeFontValues.MajorHighAnsi,
                    EastAsiaTheme = ThemeFontValues.MajorEastAsia,
                    ComplexScriptTheme = ThemeFontValues.MajorBidi
                };
        }

        /// <summary>Creates Heading 1–6 from the theme's type scale and text colour.</summary>
        internal static Style CreateHeading(int level) {
            if (level < 1 || level > 6) throw new ArgumentOutOfRangeException(nameof(level));
            string levelText = level.ToString(CultureInfo.InvariantCulture);
            OfficeThemeParagraphSpacing spacing = Theme.TextStyle.GetHeadingSpacing(level);
            var paragraphProperties = new StyleParagraphProperties(
                new KeepNext(),
                new KeepLines(),
                new SpacingBetweenLines { Before = Twips(spacing.Before.ToPoints()), After = Twips(spacing.After.ToPoints()) },
                new OutlineLevel { Val = level - 1 });
            var runProperties = new StyleRunProperties(HeadingFonts());
            if (Theme.TextStyle.HeadingBold) {
                runProperties.Append(new Bold());
                runProperties.Append(new BoldComplexScript());
            }
            runProperties.Append(new Color { Val = Hex(Theme.TextStyle.HeadingColor) });
            runProperties.Append(new FontSize { Val = HalfPoints(HeadingSize(level)) });
            runProperties.Append(new FontSizeComplexScript { Val = HalfPoints(HeadingSize(level)) });

            return new Style(
                new StyleName { Val = "heading " + levelText },
                new BasedOn { Val = "Normal" },
                new NextParagraphStyle { Val = "Normal" },
                new LinkedStyle { Val = "Heading" + levelText + "Char" },
                new UIPriority { Val = 9 },
                level == 1 ? new PrimaryStyle() : new UnhideWhenUsed(),
                paragraphProperties,
                runProperties) {
                Type = StyleValues.Paragraph,
                StyleId = "Heading" + levelText
            };
        }

        /// <summary>Creates the linked character style from the same heading font, size and colour.</summary>
        internal static Style CreateHeadingCharacter(int level) {
            Style heading = CreateHeading(level);
            return new Style(
                new StyleName { Val = "Heading " + level.ToString(CultureInfo.InvariantCulture) + " Char" },
                new BasedOn { Val = "DefaultParagraphFont" },
                new LinkedStyle { Val = heading.StyleId },
                new UIPriority { Val = 9 },
                new SemiHidden(),
                new UnhideWhenUsed(),
                heading.StyleRunProperties!.CloneNode(true)) {
                Type = StyleValues.Character,
                CustomStyle = true,
                StyleId = heading.StyleId + "Char"
            };
        }

        /// <summary>Creates the Title style: title size, heading font and text colour.</summary>
        internal static Style CreateTitle() {
            var runProperties = new StyleRunProperties(HeadingFonts());
            if (Theme.TextStyle.HeadingBold) {
                runProperties.Append(new Bold());
                runProperties.Append(new BoldComplexScript());
            }
            runProperties.Append(new Color { Val = Hex(Theme.TextStyle.TitleColor) });
            runProperties.Append(new Kern { Val = 28U });
            runProperties.Append(new FontSize { Val = HalfPoints(Theme.TypeScale.Title) });
            runProperties.Append(new FontSizeComplexScript { Val = HalfPoints(Theme.TypeScale.Title) });
            return new Style(
                new StyleName { Val = "Title" },
                new BasedOn { Val = "Normal" },
                new NextParagraphStyle { Val = "Normal" },
                new UIPriority { Val = 10 },
                new PrimaryStyle(),
                new StyleParagraphProperties(
                    new KeepNext(),
                    new SpacingBetweenLines { Before = Twips(Theme.TextStyle.TitleSpacing.Before.ToPoints()), After = Twips(Theme.TextStyle.TitleSpacing.After.ToPoints()), Line = "240", LineRule = LineSpacingRuleValues.Auto },
                    new ContextualSpacing()),
                runProperties) {
                Type = StyleValues.Paragraph,
                StyleId = "Title"
            };
        }

        /// <summary>Creates the Subtitle style: Heading 3 size in the first accent colour.</summary>
        internal static Style CreateSubtitle() {
            return new Style(
                new StyleName { Val = "Subtitle" },
                new BasedOn { Val = "Normal" },
                new NextParagraphStyle { Val = "Normal" },
                new UIPriority { Val = 11 },
                new PrimaryStyle(),
                new StyleParagraphProperties(
                    new KeepNext(),
                    new SpacingBetweenLines { Before = "0", After = Twips(Theme.Spacing.ParagraphAfter.ToPoints() * 1.5D) }),
                new StyleRunProperties(
                    new Color { Val = Hex(Theme.TextStyle.SubtitleColor) },
                    new FontSize { Val = HalfPoints(Theme.TypeScale.Heading3) },
                    new FontSizeComplexScript { Val = HalfPoints(Theme.TypeScale.Heading3) })) {
                Type = StyleValues.Paragraph,
                StyleId = "Subtitle"
            };
        }

        /// <summary>
        /// Creates the default table style: a tinted, bold, repeated header row with an
        /// accent rule beneath it, hairline rows, no vertical borders, banding, theme cell
        /// padding and body text one point smaller than Normal.
        /// </summary>
        internal static Style CreateTableStyle() {
            OfficeThemeTableStyle table = Theme.TableStyle;
            string border = Hex(table.BorderColor);
            uint borderSize = EighthPoints(table.BorderWidth.ToPoints());
            double padding = table.CellPadding.ToPoints();
            string vertical = Twips(padding);
            string horizontal = Twips(padding * TableHorizontalPaddingRatio);
            string bodySize = HalfPoints(OfficeLength.Points(Math.Max(1D, Theme.TypeScale.Body.ToPoints() - 1D)));

            var style = new Style(
                new StyleName { Val = "OfficeIMO " + Theme.Name },
                new BasedOn { Val = "TableNormal" },
                new UIPriority { Val = 40 },
                new StyleParagraphProperties(
                    new SpacingBetweenLines { Before = "0", After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto }),
                new StyleRunProperties(
                    new Color { Val = Hex(OfficeColor.Text) },
                    new FontSize { Val = bodySize },
                    new FontSizeComplexScript { Val = bodySize }),
                new StyleTableProperties(
                    new TableStyleRowBandSize { Val = 1 },
                    new TableStyleColumnBandSize { Val = 1 },
                    new TableBorders(
                        new BottomBorder { Val = BorderValues.Single, Color = border, Size = borderSize, Space = 0U },
                        new InsideHorizontalBorder { Val = BorderValues.Single, Color = border, Size = borderSize, Space = 0U }),
                    new TableCellMarginDefault(
                        new TopMargin { Width = vertical, Type = TableWidthUnitValues.Dxa },
                        new TableCellLeftMargin { Width = short.Parse(horizontal, CultureInfo.InvariantCulture), Type = TableWidthValues.Dxa },
                        new BottomMargin { Width = vertical, Type = TableWidthUnitValues.Dxa },
                        new TableCellRightMargin { Width = short.Parse(horizontal, CultureInfo.InvariantCulture), Type = TableWidthValues.Dxa }))) {
                Type = StyleValues.Table,
                StyleId = TableStyleId
            };

            var header = new TableStyleProperties(
                new StyleParagraphProperties(new KeepNext()),
                new RunPropertiesBaseStyle(
                    new Bold { Val = table.HeaderBold },
                    new BoldComplexScript { Val = table.HeaderBold },
                    new Color { Val = Hex(table.HeaderText) }),
                new TableStyleConditionalFormattingTableRowProperties(new TableHeader()),
                new TableStyleConditionalFormattingTableCellProperties(
                    new TableCellBorders(
                        new BottomBorder { Val = table.HeaderRule.IsVisible ? BorderValues.Single : BorderValues.Nil,
                            Color = Hex(table.HeaderRule.Color), Size = EighthPoints(table.HeaderRule.Width.ToPoints()), Space = 0U }),
                    new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = Hex(table.HeaderFill) })) {
                Type = TableStyleOverrideValues.FirstRow
            };
            style.Append(header);

            if (table.Banding) {
                style.Append(new TableStyleProperties(
                    new TableStyleConditionalFormattingTableCellProperties(
                        new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = Hex(table.BandFill) })) {
                    Type = TableStyleOverrideValues.Band2Horizontal
                });
            }

            return style;
        }
    }
}
