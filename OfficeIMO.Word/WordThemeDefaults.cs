using System.Globalization;
using DocumentFormat.OpenXml.Wordprocessing;
using Style = DocumentFormat.OpenXml.Wordprocessing.Style;

namespace OfficeIMO.Word {
    /// <summary>Adapts immutable shared theme tokens to native Word table styles.</summary>
    internal static class WordThemeDefaults {
        internal const string TableStyleId = "OfficeIMOModern";
        // Native layout policy retained from the canonical Word table-style builder.
        private const double TableHorizontalPaddingRatio = 1.4D;
        private static string Hex(OfficeColor color, OfficeTheme theme) => color.Resolve(theme).ToRgbHex();
        // Word retains a light page for screen themes. Keep the dark header pair,
        // while page text and unpainted/striped body rows use a readable light-page pair.
        internal static OfficeColor PageText(OfficeTheme theme) => theme.IsDark ? theme.Palette.Background : theme.Palette.Text;
        private static string HalfPoints(OfficeLength size) =>
            Math.Max(2, (int)Math.Round(size.ToPoints() * 2D, MidpointRounding.AwayFromZero)).ToString(CultureInfo.InvariantCulture);
        private static string Twips(double points) =>
            ((int)Math.Round(points * 20D, MidpointRounding.AwayFromZero)).ToString(CultureInfo.InvariantCulture);
        private static uint EighthPoints(double points) =>
            (uint)Math.Min(96, Math.Max(2, Math.Round(points * 8D, MidpointRounding.AwayFromZero)));

        /// <summary>Builds a fresh table definition for one document, without caching caller themes.</summary>
        internal static Style CreateTableStyle(OfficeTheme theme) {
            OfficeThemeTableStyle table = theme.TableStyle;
            string border = Hex(table.BorderColor, theme);
            uint borderSize = EighthPoints(table.BorderWidth.ToPoints());
            double padding = table.CellPadding.ToPoints();
            string vertical = Twips(padding);
            string horizontal = Twips(padding * TableHorizontalPaddingRatio);
            string bodySize = HalfPoints(OfficeLength.Points(Math.Max(1D, theme.TypeScale.Body.ToPoints() - 1D)));

            var style = new Style(
                new StyleName { Val = "OfficeIMO " + theme.Name },
                new BasedOn { Val = "TableNormal" },
                new UIPriority { Val = 40 },
                new StyleParagraphProperties(
                    new SpacingBetweenLines { Before = "0", After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto }),
                new StyleRunProperties(
                    new RunFonts { Ascii = theme.Fonts.Body.Family, HighAnsi = theme.Fonts.Body.Family, EastAsia = theme.Fonts.Body.Family, ComplexScript = theme.Fonts.Body.Family },
                    new Color { Val = Hex(PageText(theme), theme) },
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
                    new Color { Val = Hex(table.HeaderText, theme) }),
                new TableStyleConditionalFormattingTableRowProperties(new TableHeader()),
                new TableStyleConditionalFormattingTableCellProperties(
                    new TableCellBorders(
                        new BottomBorder { Val = table.HeaderRule.IsVisible ? BorderValues.Single : BorderValues.Nil,
                            Color = Hex(table.HeaderRule.Color, theme), Size = EighthPoints(table.HeaderRule.Width.ToPoints()), Space = 0U }),
                    new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = Hex(table.HeaderFill, theme) })) {
                Type = TableStyleOverrideValues.FirstRow
            };
            style.Append(header);
            TableBorders borders = style.StyleTableProperties!.GetFirstChild<TableBorders>()!;
            if (table.BorderMode == OfficeThemeTableBorderMode.None) {
                borders.BottomBorder!.Val = BorderValues.Nil;
                borders.InsideHorizontalBorder!.Val = BorderValues.Nil;
            } else if (table.BorderMode == OfficeThemeTableBorderMode.Grid) {
                borders.PrependChild(new TopBorder { Val = BorderValues.Single, Color = border, Size = borderSize, Space = 0U });
                borders.Append(new LeftBorder { Val = BorderValues.Single, Color = border, Size = borderSize, Space = 0U });
                borders.Append(new RightBorder { Val = BorderValues.Single, Color = border, Size = borderSize, Space = 0U });
                borders.Append(new InsideVerticalBorder { Val = BorderValues.Single, Color = border, Size = borderSize, Space = 0U });
            }


            var orderedBorders = borders.ChildElements.OrderBy(element => element.LocalName switch {
                "top" => 0, "left" => 1, "bottom" => 2, "right" => 3, "insideH" => 4, _ => 5
            }).ToList();
            borders.RemoveAllChildren();
            foreach (var element in orderedBorders) borders.Append(element);

            if (table.Banding) {
                style.Append(new TableStyleProperties(
                    new TableStyleConditionalFormattingTableCellProperties(
                        new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = Hex(theme.IsDark ? OfficeTheme.Modern.TableStyle.BandFill : table.BandFill, theme) })) {
                    Type = TableStyleOverrideValues.Band2Horizontal
                });
            }

            return style;
        }
    }
}
