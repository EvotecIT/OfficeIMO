using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Projection {
    /// <summary>Projects BIFF12 formatting collections without changing their cell-format indices.</summary>
    internal static class XlsbStylesheetProjector {
        internal static void Install(ExcelDocument document, XlsbStylesheet source) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (source == null) throw new ArgumentNullException(nameof(source));

            WorkbookStylesPart stylesPart = document.WorkbookPartRoot.WorkbookStylesPart
                ?? document.WorkbookPartRoot.AddNewPart<WorkbookStylesPart>();
            Stylesheet stylesheet = Create(source);
            stylesPart.Stylesheet = stylesheet;
            stylesheet.Save();
        }

        internal static Stylesheet Create(XlsbStylesheet source) {
            if (source == null) throw new ArgumentNullException(nameof(source));
            var stylesheet = new Stylesheet();

            if (source.NumberFormats.Count > 0) {
                var numberingFormats = new NumberingFormats();
                foreach (KeyValuePair<ushort, string> numberFormat in source.NumberFormats.OrderBy(item => item.Key)) {
                    numberingFormats.Append(new NumberingFormat {
                        NumberFormatId = numberFormat.Key,
                        FormatCode = numberFormat.Value
                    });
                }
                numberingFormats.Count = (uint)numberingFormats.Count();
                stylesheet.NumberingFormats = numberingFormats;
            }

            stylesheet.Fonts = new Fonts(source.Fonts.Select(CreateFont));
            stylesheet.Fonts.Count = (uint)stylesheet.Fonts.Count();
            stylesheet.Fills = new Fills(source.Fills.Select(CreateFill));
            stylesheet.Fills.Count = (uint)stylesheet.Fills.Count();
            stylesheet.Borders = new Borders(source.Borders.Select(CreateBorder));
            stylesheet.Borders.Count = (uint)stylesheet.Borders.Count();
            stylesheet.CellStyleFormats = new CellStyleFormats(source.CellStyleFormats.Select(format => CreateCellFormat(format, isStyleFormat: true)));
            stylesheet.CellStyleFormats.Count = (uint)stylesheet.CellStyleFormats.Count();
            stylesheet.CellFormats = new CellFormats(source.CellFormats.Select(format => CreateCellFormat(format, isStyleFormat: false)));
            stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Count();
            stylesheet.CellStyles = new CellStyles(new CellStyle {
                Name = "Normal",
                FormatId = 0U,
                BuiltinId = 0U
            }) { Count = 1U };
            stylesheet.DifferentialFormats = new DifferentialFormats { Count = 0U };
            stylesheet.TableStyles = new TableStyles {
                Count = 0U,
                DefaultTableStyle = "TableStyleMedium2",
                DefaultPivotStyle = "PivotStyleLight16"
            };

            return stylesheet;
        }

        private static Font CreateFont(XlsbFont source) {
            var font = new Font {
                FontName = new FontName { Val = source.Name },
                FontSize = new FontSize { Val = source.HeightTwips / 20D }
            };
            if (source.Weight >= 700) font.Bold = new Bold();
            if ((source.Flags & 0x0002) != 0) font.Italic = new Italic();
            if ((source.Flags & 0x0008) != 0) font.Strike = new Strike();
            if ((source.Flags & 0x0010) != 0) font.Outline = new Outline();
            if ((source.Flags & 0x0020) != 0) font.Shadow = new Shadow();
            if ((source.Flags & 0x0040) != 0) font.Condense = new Condense();
            if ((source.Flags & 0x0080) != 0) font.Extend = new Extend();
            if (ToUnderline(source.Underline) is UnderlineValues underline) {
                font.Underline = new Underline { Val = underline };
            }
            if (ToVerticalTextAlignment(source.Script) is VerticalAlignmentRunValues verticalAlignment) {
                font.VerticalTextAlignment = new VerticalTextAlignment { Val = verticalAlignment };
            }
            if (source.Color != null) font.Color = CreateColor(source.Color);
            if (source.Family != 0) font.FontFamilyNumbering = new FontFamilyNumbering { Val = source.Family };
            if (source.CharacterSet != 1) font.AddChild(new FontCharSet { Val = source.CharacterSet }, true);
            if (source.Scheme == 1) font.AddChild(new FontScheme { Val = FontSchemeValues.Major }, true);
            if (source.Scheme == 2) font.AddChild(new FontScheme { Val = FontSchemeValues.Minor }, true);
            return font;
        }

        private static Fill CreateFill(XlsbFill source) {
            // Gradient records remain preserved in the XLSB source. Until their stops are projected,
            // expose the base colors as a pattern fill instead of inventing a partial gradient.
            var pattern = new PatternFill { PatternType = ToFillPattern(source.Pattern) };
            if (source.Foreground != null) pattern.ForegroundColor = CreateForegroundColor(source.Foreground);
            if (source.Background != null) pattern.BackgroundColor = CreateBackgroundColor(source.Background);
            return new Fill(pattern);
        }

        private static Border CreateBorder(XlsbBorder source) {
            return new Border {
                LeftBorder = CreateBorderSide<LeftBorder>(source.Left),
                RightBorder = CreateBorderSide<RightBorder>(source.Right),
                TopBorder = CreateBorderSide<TopBorder>(source.Top),
                BottomBorder = CreateBorderSide<BottomBorder>(source.Bottom),
                DiagonalBorder = CreateBorderSide<DiagonalBorder>(source.Diagonal),
                DiagonalDown = source.DiagonalDown ? true : null,
                DiagonalUp = source.DiagonalUp ? true : null
            };
        }

        private static T CreateBorderSide<T>(XlsbBorderSide source) where T : BorderPropertiesType, new() {
            var side = new T();
            if (ToBorderStyle(source.Style) is BorderStyleValues style) side.Style = style;
            if (source.Color != null) side.Color = CreateColor(source.Color);
            return side;
        }

        private static CellFormat CreateCellFormat(XlsbCellFormat source, bool isStyleFormat) {
            var format = new CellFormat {
                NumberFormatId = source.NumberFormatId,
                FontId = source.FontId,
                FillId = source.FillId,
                BorderId = source.BorderId,
                ApplyNumberFormat = HasApplyFlag(source, 0) ? true : null,
                ApplyFont = HasApplyFlag(source, 1) ? true : null,
                ApplyAlignment = HasApplyFlag(source, 2) ? true : null,
                ApplyBorder = HasApplyFlag(source, 3) ? true : null,
                ApplyFill = HasApplyFlag(source, 4) ? true : null,
                ApplyProtection = HasApplyFlag(source, 5) ? true : null,
                PivotButton = source.PivotButton ? true : null,
                QuotePrefix = source.QuotePrefix ? true : null
            };
            if (!isStyleFormat) format.FormatId = source.ParentFormatId;

            bool hasAlignment = source.HorizontalAlignment != 0
                || source.VerticalAlignment != 0
                || source.TextRotation != 0
                || source.Indent != 0
                || source.WrapText
                || source.JustifyLastLine
                || source.ShrinkToFit
                || source.Merged
                || source.ReadingOrder != 0;
            if (hasAlignment || HasApplyFlag(source, 2)) {
                format.Alignment = new Alignment {
                    Horizontal = ToHorizontalAlignment(source.HorizontalAlignment),
                    Vertical = ToVerticalAlignment(source.VerticalAlignment),
                    TextRotation = source.TextRotation == 0 ? null : source.TextRotation,
                    Indent = source.Indent == 0 ? null : source.Indent,
                    WrapText = source.WrapText ? true : null,
                    JustifyLastLine = source.JustifyLastLine ? true : null,
                    ShrinkToFit = source.ShrinkToFit ? true : null,
                    MergeCell = source.Merged ? "1" : null,
                    ReadingOrder = source.ReadingOrder == 0 ? null : source.ReadingOrder
                };
            }

            if (source.Locked || source.Hidden || HasApplyFlag(source, 5)) {
                format.Protection = new Protection {
                    Locked = source.Locked,
                    Hidden = source.Hidden ? true : null
                };
            }
            return format;
        }

        private static bool HasApplyFlag(XlsbCellFormat source, int bit) => (source.ApplyFlags & (1 << bit)) != 0;

        private static DocumentFormat.OpenXml.Spreadsheet.Color CreateColor(XlsbColor source) {
            var color = new DocumentFormat.OpenXml.Spreadsheet.Color();
            ApplyColor(color, source);
            return color;
        }

        private static ForegroundColor CreateForegroundColor(XlsbColor source) {
            var color = new ForegroundColor();
            ApplyColor(color, source);
            return color;
        }

        private static BackgroundColor CreateBackgroundColor(XlsbColor source) {
            var color = new BackgroundColor();
            ApplyColor(color, source);
            return color;
        }

        private static void ApplyColor(ColorType target, XlsbColor source) {
            switch (source.Type) {
                case 0:
                    target.Auto = true;
                    break;
                case 1:
                    target.Indexed = source.Index;
                    break;
                case 2:
                    target.Rgb = $"{source.Alpha:X2}{source.Red:X2}{source.Green:X2}{source.Blue:X2}";
                    break;
                case 3:
                    target.Theme = source.Index;
                    break;
            }

            if (source.Tint != 0) {
                target.Tint = source.Tint / (source.Tint < 0 ? 32768D : 32767D);
            }
        }

        private static UnderlineValues? ToUnderline(byte value) {
            return value switch {
                1 => UnderlineValues.Single,
                2 => UnderlineValues.Double,
                0x21 => UnderlineValues.SingleAccounting,
                0x22 => UnderlineValues.DoubleAccounting,
                _ => null
            };
        }

        private static VerticalAlignmentRunValues? ToVerticalTextAlignment(ushort value) {
            return value switch {
                1 => VerticalAlignmentRunValues.Superscript,
                2 => VerticalAlignmentRunValues.Subscript,
                _ => null
            };
        }

        private static PatternValues ToFillPattern(uint value) {
            return value switch {
                1 => PatternValues.Solid,
                2 => PatternValues.MediumGray,
                3 => PatternValues.DarkGray,
                4 => PatternValues.LightGray,
                5 => PatternValues.DarkHorizontal,
                6 => PatternValues.DarkVertical,
                7 => PatternValues.DarkDown,
                8 => PatternValues.DarkUp,
                9 => PatternValues.DarkGrid,
                10 => PatternValues.DarkTrellis,
                11 => PatternValues.LightHorizontal,
                12 => PatternValues.LightVertical,
                13 => PatternValues.LightDown,
                14 => PatternValues.LightUp,
                15 => PatternValues.LightGrid,
                16 => PatternValues.LightTrellis,
                17 => PatternValues.Gray125,
                18 => PatternValues.Gray0625,
                _ => PatternValues.None
            };
        }

        private static BorderStyleValues? ToBorderStyle(byte value) {
            return value switch {
                1 => BorderStyleValues.Thin,
                2 => BorderStyleValues.Medium,
                3 => BorderStyleValues.Dashed,
                4 => BorderStyleValues.Dotted,
                5 => BorderStyleValues.Thick,
                6 => BorderStyleValues.Double,
                7 => BorderStyleValues.Hair,
                8 => BorderStyleValues.MediumDashed,
                9 => BorderStyleValues.DashDot,
                10 => BorderStyleValues.MediumDashDot,
                11 => BorderStyleValues.DashDotDot,
                12 => BorderStyleValues.MediumDashDotDot,
                13 => BorderStyleValues.SlantDashDot,
                _ => null
            };
        }

        private static HorizontalAlignmentValues? ToHorizontalAlignment(byte value) {
            return value switch {
                1 => HorizontalAlignmentValues.Left,
                2 => HorizontalAlignmentValues.Center,
                3 => HorizontalAlignmentValues.Right,
                4 => HorizontalAlignmentValues.Fill,
                5 => HorizontalAlignmentValues.Justify,
                6 => HorizontalAlignmentValues.CenterContinuous,
                7 => HorizontalAlignmentValues.Distributed,
                _ => null
            };
        }

        private static VerticalAlignmentValues? ToVerticalAlignment(byte value) {
            return value switch {
                0 => VerticalAlignmentValues.Top,
                1 => VerticalAlignmentValues.Center,
                2 => VerticalAlignmentValues.Bottom,
                3 => VerticalAlignmentValues.Justify,
                4 => VerticalAlignmentValues.Distributed,
                _ => null
            };
        }
    }
}
