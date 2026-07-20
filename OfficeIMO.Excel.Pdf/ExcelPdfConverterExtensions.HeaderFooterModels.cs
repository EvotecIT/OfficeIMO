using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Excel.Pdf {
    public static partial class ExcelPdfConverterExtensions {
        private sealed class HeaderFooterZone {
            internal static readonly HeaderFooterZone Empty = new HeaderFooterZone(null, null);

            internal HeaderFooterZone(string? text, HeaderFooterLineStyle? style) {
                Text = text;
                Style = style;
            }

            internal string? Text { get; }

            internal HeaderFooterLineStyle? Style { get; }
        }

        private sealed class HeaderFooterZones {
            internal HeaderFooterZones(string? left, string? center, string? right, HeaderFooterLineStyle? style) {
                Left = left;
                Center = center;
                Right = right;
                Style = style;
            }

            internal string? Left { get; }

            internal string? Center { get; }

            internal string? Right { get; }

            internal HeaderFooterLineStyle? Style { get; }
        }

        private sealed class HeaderFooterLineStyle {
            internal double? FontSize { get; set; }

            internal PdfCore.PdfColor? Color { get; set; }

            internal PdfCore.PdfStandardFont? FontFamily { get; set; }

            internal string? FontFamilyName { get; set; }

            internal bool Bold { get; set; }

            internal bool Italic { get; set; }

            internal PdfCore.PdfStandardFont? Font {
                get {
                    if (!FontFamily.HasValue && !Bold && !Italic) {
                        return null;
                    }

                    PdfCore.PdfStandardFont family = FontFamily ?? PdfCore.PdfStandardFont.Helvetica;
                    return PdfCore.PdfStandardFontMapper.GetStyledFont(family, Bold, Italic);
                }
            }

            internal bool HasAnyStyle {
                get {
                    PdfCore.PdfStandardFont? font = Font;
                    return FontSize.HasValue
                           || Color.HasValue
                           || !string.IsNullOrWhiteSpace(FontFamilyName)
                           || (font.HasValue && font.Value != PdfCore.PdfStandardFont.Helvetica);
                }
            }

            internal static bool Equals(HeaderFooterLineStyle? left, HeaderFooterLineStyle? right) {
                if (ReferenceEquals(left, right)) {
                    return true;
                }

                if (left == null || right == null) {
                    return false;
                }

                return Nullable.Equals(left.FontSize, right.FontSize)
                       && Nullable.Equals(left.Color, right.Color)
                       && Nullable.Equals(left.Font, right.Font)
                       && string.Equals(left.FontFamilyName, right.FontFamilyName, StringComparison.OrdinalIgnoreCase);
            }
        }

    }
}
