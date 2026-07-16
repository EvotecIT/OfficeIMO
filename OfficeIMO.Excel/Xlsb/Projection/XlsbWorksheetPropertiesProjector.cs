using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Projection {
    /// <summary>Projects and compares worksheet-level properties stored in BrtWsProp.</summary>
    internal static class XlsbWorksheetPropertiesProjector {
        internal static void Apply(ExcelSheet sheet, XlsbWorksheetProperties? source) {
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            if (source == null) return;
            Worksheet worksheet = sheet.WorksheetPart.Worksheet
                ?? throw new InvalidDataException($"Worksheet '{sheet.Name}' has no worksheet root.");
            worksheet.PrependChild(Create(source));
        }

        internal static void ValidateUnchanged(ExcelSheet sheet, XlsbWorksheetProperties? source) {
            Worksheet worksheet = sheet.WorksheetPart.Worksheet
                ?? throw new InvalidDataException($"Worksheet '{sheet.Name}' has no worksheet root.");
            SheetProperties[] actual = worksheet.Elements<SheetProperties>().ToArray();
            SheetProperties? expected = source == null ? null : Create(source);
            if (actual.Length > 1
                || (expected == null && actual.Length != 0)
                || (expected != null
                    && (actual.Length != 1
                        || !string.Equals(actual[0].OuterXml, expected.OuterXml, StringComparison.Ordinal)))) {
                throw new NotSupportedException($"Native XLSB rewriting preserves but cannot modify worksheet properties on worksheet '{sheet.Name}'. Save as .xlsx to retain that change.");
            }
        }

        private static SheetProperties Create(XlsbWorksheetProperties source) {
            var properties = new SheetProperties {
                Published = source.Published,
                SyncHorizontal = source.SynchronizeHorizontal,
                SyncVertical = source.SynchronizeVertical,
                TransitionEvaluation = source.TransitionEvaluation,
                TransitionEntry = source.TransitionEntry,
                FilterMode = source.FilterMode,
                EnableFormatConditionsCalculation = source.CalculateConditionalFormatting
            };
            if (!string.IsNullOrEmpty(source.CodeName)) properties.CodeName = source.CodeName;
            if ((source.SynchronizeHorizontal || source.SynchronizeVertical)
                && source.SynchronizedRow != uint.MaxValue
                && source.SynchronizedColumn != uint.MaxValue) {
                properties.SyncReference = A1.CellReference(
                    checked((int)source.SynchronizedRow + 1),
                    checked((int)source.SynchronizedColumn + 1));
            }
            if (source.TabColor?.Type is > 0 and <= 3) {
                var color = new TabColor();
                ApplyColor(color, source.TabColor);
                properties.TabColor = color;
            }
            properties.Append(new OutlineProperties {
                ApplyStyles = source.ApplyOutlineStyles,
                SummaryBelow = source.SummaryRowsBelow,
                SummaryRight = source.SummaryColumnsRight,
                ShowOutlineSymbols = source.ShowOutlineSymbols
            });
            properties.Append(new PageSetupProperties {
                AutoPageBreaks = source.ShowAutomaticPageBreaks,
                FitToPage = source.FitToPage
            });
            return properties;
        }

        private static void ApplyColor(ColorType target, XlsbColor source) {
            if (source.Type == 1) target.Indexed = source.Index;
            else if (source.Type == 2) target.Rgb = $"{source.Alpha:X2}{source.Red:X2}{source.Green:X2}{source.Blue:X2}";
            else if (source.Type == 3) target.Theme = source.Index;
            if (source.Tint != 0) {
                target.Tint = source.Tint / (source.Tint < 0 ? 32768D : 32767D);
            }
        }
    }
}
