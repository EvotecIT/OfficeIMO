using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Projection {
    internal static class LegacyXlsConditionalFormattingProjector {
        internal static void Project(ExcelSheet sheet, LegacyXlsConditionalFormatting conditionalFormatting) {
            string range = string.Join(" ", conditionalFormatting.Ranges);
            if (string.IsNullOrWhiteSpace(range) || string.IsNullOrWhiteSpace(conditionalFormatting.Formula1)) {
                return;
            }

            switch (conditionalFormatting.Type) {
                case LegacyXlsConditionalFormattingType.CellIs:
                    if (!conditionalFormatting.Operator.HasValue) {
                        return;
                    }

                    sheet.AddConditionalRule(
                        range,
                        ToOperator(conditionalFormatting.Operator.Value),
                        conditionalFormatting.Formula1,
                        conditionalFormatting.Formula2,
                        conditionalFormatting.StopIfTrue,
                        conditionalFormatting.Priority);
                    break;
                case LegacyXlsConditionalFormattingType.Formula:
                    sheet.AddConditionalFormulaRule(range, conditionalFormatting.Formula1, conditionalFormatting.StopIfTrue, conditionalFormatting.Priority);
                    break;
            }
        }

        private static ConditionalFormattingOperatorValues ToOperator(LegacyXlsConditionalFormattingOperator @operator) {
            return @operator switch {
                LegacyXlsConditionalFormattingOperator.Between => ConditionalFormattingOperatorValues.Between,
                LegacyXlsConditionalFormattingOperator.NotBetween => ConditionalFormattingOperatorValues.NotBetween,
                LegacyXlsConditionalFormattingOperator.Equal => ConditionalFormattingOperatorValues.Equal,
                LegacyXlsConditionalFormattingOperator.NotEqual => ConditionalFormattingOperatorValues.NotEqual,
                LegacyXlsConditionalFormattingOperator.GreaterThan => ConditionalFormattingOperatorValues.GreaterThan,
                LegacyXlsConditionalFormattingOperator.LessThan => ConditionalFormattingOperatorValues.LessThan,
                LegacyXlsConditionalFormattingOperator.GreaterThanOrEqual => ConditionalFormattingOperatorValues.GreaterThanOrEqual,
                _ => ConditionalFormattingOperatorValues.LessThanOrEqual
            };
        }
    }
}
