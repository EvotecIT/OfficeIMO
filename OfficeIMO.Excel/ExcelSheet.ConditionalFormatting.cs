using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using SixLabors.Fonts;
using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using SixLaborsColor = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        public void AddConditionalRule(string range, ConditionalFormattingOperatorValues @operator, string formula1, string? formula2 = null) {
            if (string.IsNullOrEmpty(range)) {
                throw new ArgumentNullException(nameof(range));
            }

            WriteLock(() => {
                Worksheet worksheet = _worksheetPart.Worksheet;

                int priority = 1;
                var existingRules = worksheet.Descendants<ConditionalFormattingRule>();
                if (existingRules.Any()) {
                    priority = existingRules.Max(r => r.Priority?.Value ?? 0) + 1;
                }

                ConditionalFormatting conditionalFormatting = new ConditionalFormatting {
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = range }
                };

                ConditionalFormattingRule rule = new ConditionalFormattingRule {
                    Type = ConditionalFormatValues.CellIs,
                    Operator = @operator,
                    Priority = priority
                };

                rule.Append(new Formula(formula1));
                if (formula2 != null) {
                    rule.Append(new Formula(formula2));
                }

                conditionalFormatting.Append(rule);
                
                // Insert ConditionalFormatting after AutoFilter but before TableParts
                var autoFilter = worksheet.Elements<AutoFilter>().FirstOrDefault();
                var tableParts = worksheet.Elements<TableParts>().FirstOrDefault();
                
                if (tableParts != null) {
                    worksheet.InsertBefore(conditionalFormatting, tableParts);
                } else if (autoFilter != null) {
                    worksheet.InsertAfter(conditionalFormatting, autoFilter);
                } else {
                    var sheetData = worksheet.GetFirstChild<SheetData>();
                    if (sheetData != null) {
                        worksheet.InsertAfter(conditionalFormatting, sheetData);
                    } else {
                        worksheet.Append(conditionalFormatting);
                    }
                }
                
                worksheet.Save();
            });
        }

        private static string ConvertColor(SixLaborsColor color) {
            return "FF" + color.ToHexColor();
        }

        public void AddConditionalColorScale(string range, SixLaborsColor startColor, SixLaborsColor endColor) {
            AddConditionalColorScale(range, ConvertColor(startColor), ConvertColor(endColor));
        }

        public void AddConditionalColorScale(string range, string startColor, string endColor) {
            if (string.IsNullOrEmpty(range)) {
                throw new ArgumentNullException(nameof(range));
            }

            WriteLock(() => {
                Worksheet worksheet = _worksheetPart.Worksheet;

                int priority = 1;
                var existingRules = worksheet.Descendants<ConditionalFormattingRule>();
                if (existingRules.Any()) {
                    priority = existingRules.Max(r => r.Priority?.Value ?? 0) + 1;
                }

                ConditionalFormatting conditionalFormatting = new ConditionalFormatting {
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = range }
                };

                ConditionalFormattingRule rule = new ConditionalFormattingRule {
                    Type = ConditionalFormatValues.ColorScale,
                    Priority = priority
                };

                ColorScale colorScale = new ColorScale();
                colorScale.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Min });
                colorScale.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Max });
                colorScale.Append(new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = startColor });
                colorScale.Append(new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = endColor });
                rule.Append(colorScale);

                conditionalFormatting.Append(rule);
                
                // Insert ConditionalFormatting after AutoFilter but before TableParts
                var autoFilter = worksheet.Elements<AutoFilter>().FirstOrDefault();
                var tableParts = worksheet.Elements<TableParts>().FirstOrDefault();
                
                if (tableParts != null) {
                    worksheet.InsertBefore(conditionalFormatting, tableParts);
                } else if (autoFilter != null) {
                    worksheet.InsertAfter(conditionalFormatting, autoFilter);
                } else {
                    var sheetData = worksheet.GetFirstChild<SheetData>();
                    if (sheetData != null) {
                        worksheet.InsertAfter(conditionalFormatting, sheetData);
                    } else {
                        worksheet.Append(conditionalFormatting);
                    }
                }
                
                worksheet.Save();
            });
        }

        public void AddConditionalDataBar(string range, SixLaborsColor color) {
            AddConditionalDataBar(range, ConvertColor(color));
        }

        public void AddConditionalDataBar(string range, string color) {
            if (string.IsNullOrEmpty(range)) {
                throw new ArgumentNullException(nameof(range));
            }

            WriteLock(() => {
                Worksheet worksheet = _worksheetPart.Worksheet;

                int priority = 1;
                var existingRules = worksheet.Descendants<ConditionalFormattingRule>();
                if (existingRules.Any()) {
                    priority = existingRules.Max(r => r.Priority?.Value ?? 0) + 1;
                }

                ConditionalFormatting conditionalFormatting = new ConditionalFormatting {
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = range }
                };

                ConditionalFormattingRule rule = new ConditionalFormattingRule {
                    Type = ConditionalFormatValues.DataBar,
                    Priority = priority
                };

                DataBar dataBar = new DataBar { ShowValue = true };
                dataBar.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Min });
                dataBar.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Max });
                dataBar.Append(new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = color });
                rule.Append(dataBar);

                conditionalFormatting.Append(rule);
                
                // Insert ConditionalFormatting after AutoFilter but before TableParts
                var autoFilter = worksheet.Elements<AutoFilter>().FirstOrDefault();
                var tableParts = worksheet.Elements<TableParts>().FirstOrDefault();
                
                if (tableParts != null) {
                    worksheet.InsertBefore(conditionalFormatting, tableParts);
                } else if (autoFilter != null) {
                    worksheet.InsertAfter(conditionalFormatting, autoFilter);
                } else {
                    var sheetData = worksheet.GetFirstChild<SheetData>();
                    if (sheetData != null) {
                        worksheet.InsertAfter(conditionalFormatting, sheetData);
                    } else {
                        worksheet.Append(conditionalFormatting);
                    }
                }
                
                worksheet.Save();
            });
        }

    }
}

