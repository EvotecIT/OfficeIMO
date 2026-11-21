using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>
    /// Encodes worksheet conditionalFormatting elements as CONDITIONALFORMATTING collections (MS-XLSB 2.4.34, 2.4.23,
    /// 2.4.28, 2.4.43, 2.4.91, 2.4.334, 2.4.337). Office 2010 extension rules (x14) are refused.
    /// </summary>
    internal static class XlsbWorksheetConditionalFormattingWriter {
        internal const int BrtBeginConditionalFormatting = 461;
        internal const int BrtEndConditionalFormatting = 462;
        internal const int BrtBeginCFRule = 463;
        internal const int BrtEndCFRule = 464;
        internal const int BrtBeginIconSet = 465;
        internal const int BrtEndIconSet = 466;
        internal const int BrtBeginDatabar = 467;
        internal const int BrtEndDatabar = 468;
        internal const int BrtBeginColorScale = 469;
        internal const int BrtEndColorScale = 470;
        internal const int BrtCFVO = 471;
        internal const int BrtColor = 564;
        private const int MaximumRanges = 8192;
        private const uint NoDifferentialFormat = 0xFFFFFFFFU;

        // CFType (MS-XLSB 2.5.18).
        internal const uint TypeCellIs = 1, TypeExpression = 2, TypeGradient = 3, TypeDataBar = 4, TypeFilter = 5, TypeMultiState = 6;

        // CFTemp (MS-XLSB 2.5.16).
        internal const uint TemplateExpression = 0x00, TemplateFormula = 0x01, TemplateGradient = 0x02, TemplateDataBar = 0x03,
            TemplateMultiState = 0x04, TemplateFilter = 0x05, TemplateUniqueValues = 0x07, TemplateContainsText = 0x08,
            TemplateContainsBlanks = 0x09, TemplateContainsNoBlanks = 0x0A, TemplateContainsErrors = 0x0B, TemplateContainsNoErrors = 0x0C,
            TemplateAboveAverage = 0x19, TemplateBelowAverage = 0x1A, TemplateDuplicateValues = 0x1B,
            TemplateEqualAboveAverage = 0x1D, TemplateEqualBelowAverage = 0x1E;

        /// <summary>CFOper (MS-XLSB 2.5.15) values 1 through 8, in order.</summary>
        internal static readonly ConditionalFormattingOperatorValues[] CellOperators = {
            ConditionalFormattingOperatorValues.Between,
            ConditionalFormattingOperatorValues.NotBetween,
            ConditionalFormattingOperatorValues.Equal,
            ConditionalFormattingOperatorValues.NotEqual,
            ConditionalFormattingOperatorValues.GreaterThan,
            ConditionalFormattingOperatorValues.LessThan,
            ConditionalFormattingOperatorValues.GreaterThanOrEqual,
            ConditionalFormattingOperatorValues.LessThanOrEqual,
        };

        /// <summary>CFTextOper (MS-XLSB 2.5.17) values 0 through 3 with their SpreadsheetML rule type and operator.</summary>
        internal static readonly (ConditionalFormatValues Type, ConditionalFormattingOperatorValues Operator)[] TextOperators = {
            (ConditionalFormatValues.ContainsText, ConditionalFormattingOperatorValues.ContainsText),
            (ConditionalFormatValues.NotContainsText, ConditionalFormattingOperatorValues.NotContains),
            (ConditionalFormatValues.BeginsWith, ConditionalFormattingOperatorValues.BeginsWith),
            (ConditionalFormatValues.EndsWith, ConditionalFormattingOperatorValues.EndsWith),
        };

        /// <summary>Time periods with their CFTemp template (0x0F through 0x18) and CFDateOper (MS-XLSB 2.5.12) parameter.</summary>
        internal static readonly (TimePeriodValues Period, uint Template, uint Parameter)[] TimePeriods = {
            (TimePeriodValues.Today, 0x0F, 0),
            (TimePeriodValues.Tomorrow, 0x10, 6),
            (TimePeriodValues.Yesterday, 0x11, 1),
            (TimePeriodValues.Last7Days, 0x12, 2),
            (TimePeriodValues.LastMonth, 0x13, 5),
            (TimePeriodValues.NextMonth, 0x14, 8),
            (TimePeriodValues.ThisWeek, 0x15, 3),
            (TimePeriodValues.NextWeek, 0x16, 7),
            (TimePeriodValues.LastWeek, 0x17, 4),
            (TimePeriodValues.ThisMonth, 0x18, 9),
        };

        /// <summary>CFVOtype (MS-XLSB 2.5.19) values with their SpreadsheetML cfvo type.</summary>
        internal static readonly (ConditionalFormatValueObjectValues Type, uint Value)[] ValueObjectTypes = {
            (ConditionalFormatValueObjectValues.Number, 1),
            (ConditionalFormatValueObjectValues.Min, 2),
            (ConditionalFormatValueObjectValues.Max, 3),
            (ConditionalFormatValueObjectValues.Percent, 4),
            (ConditionalFormatValueObjectValues.Percentile, 5),
            (ConditionalFormatValueObjectValues.Formula, 7),
        };

        /// <summary>KPISets (MS-XLSB 2.5.86) in value order, starting at 0.</summary>
        internal static readonly IconSetValues[] IconSets = {
            IconSetValues.ThreeArrows, IconSetValues.ThreeArrowsGray, IconSetValues.ThreeFlags, IconSetValues.ThreeTrafficLights1,
            IconSetValues.ThreeTrafficLights2, IconSetValues.ThreeSigns, IconSetValues.ThreeSymbols, IconSetValues.ThreeSymbols2,
            IconSetValues.FourArrows, IconSetValues.FourArrowsGray, IconSetValues.FourRedToBlack, IconSetValues.FourRating,
            IconSetValues.FourTrafficLights, IconSetValues.FiveArrows, IconSetValues.FiveArrowsGray, IconSetValues.FiveRating,
            IconSetValues.FiveQuarters,
        };

        internal static void Append(XlsbWorksheetFeatureRecords records, XlsbWorksheetFeatureContext context) =>
            records.ConditionalFormatting.AddRange(Create(context));

        internal static IReadOnlyList<XlsbGeneratedRecord> Create(XlsbWorksheetFeatureContext context) {
            ConditionalFormatting[] collections = context.Worksheet.Elements<ConditionalFormatting>().ToArray();
            if (collections.Length == 0) return Array.Empty<XlsbGeneratedRecord>();
            string sheetName = context.Sheet.Name;
            uint differentialFormatCount = (uint)(context.Document.WorkbookPartRoot.WorkbookStylesPart?.Stylesheet?.DifferentialFormats?
                .Elements<DifferentialFormat>().Count() ?? 0);
            var priorities = new HashSet<int>();
            var result = new List<XlsbGeneratedRecord>();
            foreach (ConditionalFormatting collection in collections) {
                if (XlsbOpenXmlAttributeValidator.FindUnsupported(collection, new[] { "sqref", "pivot" }) != null
                    || collection.ChildElements.Any(child => child is not ConditionalFormattingRule)) {
                    throw Refuse(sheetName, "Office 2010 extension or unknown conditional formatting content");
                }
                ConditionalFormattingRule[] rules = collection.Elements<ConditionalFormattingRule>().ToArray();
                if (rules.Length == 0) continue;
                using (var begin = new MemoryStream()) {
                    XlsbPayloadWriter.WriteUInt32(begin, (uint)rules.Length);
                    XlsbPayloadWriter.WriteUInt32(begin, collection.Pivot?.Value == true ? 1U : 0U);
                    XlsbPayloadWriter.WriteSqRfX(begin, collection.SequenceOfReferences?.InnerText, sheetName, "conditional formatting", MaximumRanges);
                    result.Add(new XlsbGeneratedRecord(BrtBeginConditionalFormatting, begin.ToArray()));
                }
                foreach (ConditionalFormattingRule rule in rules) {
                    int priority = rule.Priority?.Value ?? 0;
                    if (priority <= 0 || !priorities.Add(priority)) {
                        throw Refuse(sheetName, $"a rule priority of {priority}; priorities must be positive and unique on the worksheet");
                    }
                    AppendRule(result, rule, context, differentialFormatCount);
                }
                result.Add(new XlsbGeneratedRecord(BrtEndConditionalFormatting, Array.Empty<byte>()));
            }
            return result;
        }

        private static void AppendRule(List<XlsbGeneratedRecord> result, ConditionalFormattingRule rule, XlsbWorksheetFeatureContext context,
            uint differentialFormatCount) {
            string sheetName = context.Sheet.Name;
            if (rule.ChildElements.Any(child => child is not Formula and not ColorScale and not DataBar and not IconSet)) {
                throw Refuse(sheetName, "a rule with Office 2010 extensions (x14), which native XLSB generation does not write");
            }
            ConditionalFormatValues type = rule.Type?.Value
                ?? throw Refuse(sheetName, "a rule without a type");
            string[] formulas = rule.Elements<Formula>().Select(formula => formula.Text ?? string.Empty).ToArray();
            var allowed = new List<string> { "type", "priority", "stopIfTrue", "dxfId" };
            uint iType, template, parameter = 0;
            bool above = false, bottom = false, percent = false;
            string? text = null;
            int formulaCount;
            OpenXmlCompositeElement? visual = null;

            if (type == ConditionalFormatValues.CellIs) {
                ConditionalFormattingOperatorValues op = rule.Operator?.Value ?? throw Refuse(sheetName, "a cellIs rule without an operator");
                int index = Array.IndexOf(CellOperators, op);
                if (index < 0) throw Refuse(sheetName, $"the cellIs operator '{rule.Operator!.InnerText}'");
                (iType, template, parameter) = (TypeCellIs, TemplateExpression, (uint)index + 1);
                formulaCount = op == ConditionalFormattingOperatorValues.Between || op == ConditionalFormattingOperatorValues.NotBetween ? 2 : 1;
                allowed.Add("operator");
            } else if (type == ConditionalFormatValues.Expression) {
                (iType, template, formulaCount) = (TypeExpression, TemplateFormula, 1);
            } else if (TextOperators.Any(item => item.Type == type)) {
                int index = Array.FindIndex(TextOperators, item => item.Type == type);
                if (rule.Operator != null && rule.Operator.Value != TextOperators[index].Operator) {
                    throw Refuse(sheetName, $"a {rule.Type!.InnerText} rule with operator '{rule.Operator.InnerText}'");
                }
                text = rule.Text?.Value;
                if (string.IsNullOrEmpty(text) || text!.Length > 255) throw Refuse(sheetName, "a text rule whose text is empty or longer than 255 characters");
                (iType, template, parameter, formulaCount) = (TypeExpression, TemplateContainsText, (uint)index, 1);
                allowed.Add("operator");
                allowed.Add("text");
            } else if (type == ConditionalFormatValues.ContainsBlanks) {
                (iType, template, formulaCount) = (TypeExpression, TemplateContainsBlanks, 1);
            } else if (type == ConditionalFormatValues.NotContainsBlanks) {
                (iType, template, formulaCount) = (TypeExpression, TemplateContainsNoBlanks, 1);
            } else if (type == ConditionalFormatValues.ContainsErrors) {
                (iType, template, formulaCount) = (TypeExpression, TemplateContainsErrors, 1);
            } else if (type == ConditionalFormatValues.NotContainsErrors) {
                (iType, template, formulaCount) = (TypeExpression, TemplateContainsNoErrors, 1);
            } else if (type == ConditionalFormatValues.TimePeriod) {
                TimePeriodValues period = rule.TimePeriod?.Value ?? throw Refuse(sheetName, "a timePeriod rule without a period");
                var entry = TimePeriods.First(item => item.Period == period);
                (iType, template, parameter, formulaCount) = (TypeExpression, entry.Template, entry.Parameter, 1);
                allowed.Add("timePeriod");
            } else if (type == ConditionalFormatValues.UniqueValues) {
                (iType, template, formulaCount) = (TypeExpression, TemplateUniqueValues, 0);
            } else if (type == ConditionalFormatValues.DuplicateValues) {
                (iType, template, formulaCount) = (TypeExpression, TemplateDuplicateValues, 0);
            } else if (type == ConditionalFormatValues.AboveAverage) {
                above = rule.AboveAverage?.Value ?? true;
                bool equal = rule.EqualAverage?.Value ?? false;
                int deviations = rule.StdDev?.Value ?? 0;
                if (deviations < 0 || deviations > 3) throw Refuse(sheetName, $"an average rule with {deviations} standard deviations; XLSB allows 0 to 3");
                if (equal && deviations != 0) {
                    throw Refuse(sheetName, "an average rule that combines equalAverage with standard deviations, which XLSB cannot express");
                }
                template = equal
                    ? above ? TemplateEqualAboveAverage : TemplateEqualBelowAverage
                    : above ? TemplateAboveAverage : TemplateBelowAverage;
                (iType, parameter, formulaCount) = (TypeExpression, (uint)deviations, 0);
                allowed.AddRange(new[] { "aboveAverage", "equalAverage", "stdDev" });
            } else if (type == ConditionalFormatValues.Top10) {
                uint rank = rule.Rank?.Value ?? 10U;
                percent = rule.Percent?.Value ?? false;
                bottom = rule.Bottom?.Value ?? false;
                if (percent ? rank > 100U : rank < 1U || rank > 1000U) {
                    throw Refuse(sheetName, $"a top/bottom rank of {rank}; XLSB allows 0 to 100 percent or 1 to 1000 items");
                }
                (iType, template, parameter, formulaCount) = (TypeFilter, TemplateFilter, rank, 0);
                allowed.AddRange(new[] { "rank", "percent", "bottom" });
            } else if (type == ConditionalFormatValues.ColorScale) {
                (iType, template, formulaCount) = (TypeGradient, TemplateGradient, 0);
                visual = rule.GetFirstChild<ColorScale>();
            } else if (type == ConditionalFormatValues.DataBar) {
                (iType, template, formulaCount) = (TypeDataBar, TemplateDataBar, 0);
                visual = rule.GetFirstChild<DataBar>();
            } else if (type == ConditionalFormatValues.IconSet) {
                (iType, template, formulaCount) = (TypeMultiState, TemplateMultiState, 0);
                visual = rule.GetFirstChild<IconSet>();
            } else {
                throw Refuse(sheetName, $"the rule type '{rule.Type!.InnerText}'");
            }

            if (XlsbOpenXmlAttributeValidator.FindUnsupported(rule, allowed) is OpenXmlAttribute attribute) {
                throw Refuse(sheetName, $"the '{attribute.LocalName}' attribute on a {rule.Type!.InnerText} rule");
            }
            bool isVisual = iType is TypeGradient or TypeDataBar or TypeMultiState;
            int visualChildren = rule.ChildElements.Count(child => child is ColorScale or DataBar or IconSet);
            if (isVisual ? visual == null || visualChildren != 1 : visualChildren != 0) {
                throw Refuse(sheetName, $"a {rule.Type!.InnerText} rule with mismatched color scale, data bar or icon set content");
            }
            if (formulas.Length != formulaCount) {
                throw Refuse(sheetName, $"a {rule.Type!.InnerText} rule with {formulas.Length} formulas; XLSB stores {formulaCount}");
            }
            bool stopIfTrue = rule.StopIfTrue?.Value ?? false;
            uint differentialFormat = rule.FormatId?.Value ?? NoDifferentialFormat;
            if (isVisual) {
                if (stopIfTrue) throw Refuse(sheetName, $"stopIfTrue on a {rule.Type!.InnerText} rule");
                if (rule.FormatId != null) throw Refuse(sheetName, $"a differential format on a {rule.Type!.InnerText} rule");
            } else if (differentialFormat != NoDifferentialFormat && differentialFormat >= differentialFormatCount) {
                throw Refuse(sheetName, $"a rule that refers to missing differential format {differentialFormat}");
            }

            (int Row, int Column) anchor = FindAnchor(context.Sheet.Name, ((ConditionalFormatting)rule.Parent!).SequenceOfReferences?.InnerText);
            byte[][] encoded = formulas
                .Select(formula => EncodeRuleFormula(context, formula, anchor))
                .ToArray();
            using (var payload = new MemoryStream()) {
                XlsbPayloadWriter.WriteUInt32(payload, iType);
                XlsbPayloadWriter.WriteUInt32(payload, template);
                XlsbPayloadWriter.WriteUInt32(payload, differentialFormat);
                XlsbPayloadWriter.WriteInt32(payload, rule.Priority!.Value);
                XlsbPayloadWriter.WriteUInt32(payload, parameter);
                XlsbPayloadWriter.WriteUInt32(payload, 0); // reserved1
                XlsbPayloadWriter.WriteUInt32(payload, 0); // reserved2
                ushort flags = 0;
                if (stopIfTrue) flags |= 0x0002;
                if (above) flags |= 0x0004;
                if (bottom) flags |= 0x0008;
                if (percent) flags |= 0x0010;
                XlsbPayloadWriter.WriteUInt16(payload, flags);
                for (int index = 0; index < 3; index++) {
                    // cbFmla1..3 only signal presence; each stores the rgce length of its formula.
                    XlsbPayloadWriter.WriteUInt32(payload, index < encoded.Length ? BitConverter.ToUInt32(encoded[index], 0) : 0U);
                }
                XlsbPayloadWriter.WriteNullableWideString(payload, text);
                foreach (byte[] formula in encoded) payload.Write(formula, 0, formula.Length);
                result.Add(new XlsbGeneratedRecord(BrtBeginCFRule, payload.ToArray()));
            }

            switch (visual) {
                case ColorScale colorScale: AppendColorScale(result, colorScale, context); break;
                case DataBar dataBar: AppendDataBar(result, dataBar, context); break;
                case IconSet iconSet: AppendIconSet(result, iconSet, context); break;
            }
            result.Add(new XlsbGeneratedRecord(BrtEndCFRule, Array.Empty<byte>()));
        }

        private static void AppendColorScale(List<XlsbGeneratedRecord> result, ColorScale colorScale, XlsbWorksheetFeatureContext context) {
            ConditionalFormatValueObject[] values = colorScale.Elements<ConditionalFormatValueObject>().ToArray();
            DocumentFormat.OpenXml.Spreadsheet.Color[] colors = colorScale.Elements<DocumentFormat.OpenXml.Spreadsheet.Color>().ToArray();
            if (XlsbOpenXmlAttributeValidator.FindUnsupported(colorScale, Array.Empty<string>()) != null || values.Length + colors.Length != colorScale.ChildElements.Count
                || values.Length is < 2 or > 3 || colors.Length != values.Length) {
                throw Refuse(context.Sheet.Name, "a color scale that is not two or three value objects with one color each");
            }
            result.Add(new XlsbGeneratedRecord(BrtBeginColorScale, Array.Empty<byte>()));
            foreach (ConditionalFormatValueObject value in values) result.Add(CreateValueObject(value, iconSet: false, context));
            foreach (DocumentFormat.OpenXml.Spreadsheet.Color color in colors) result.Add(CreateColor(color, context.Sheet.Name));
            result.Add(new XlsbGeneratedRecord(BrtEndColorScale, Array.Empty<byte>()));
        }

        private static void AppendDataBar(List<XlsbGeneratedRecord> result, DataBar dataBar, XlsbWorksheetFeatureContext context) {
            ConditionalFormatValueObject[] values = dataBar.Elements<ConditionalFormatValueObject>().ToArray();
            DocumentFormat.OpenXml.Spreadsheet.Color[] colors = dataBar.Elements<DocumentFormat.OpenXml.Spreadsheet.Color>().ToArray();
            if (XlsbOpenXmlAttributeValidator.FindUnsupported(dataBar, new[] { "minLength", "maxLength", "showValue" }) != null
                || values.Length != 2 || colors.Length != 1 || dataBar.ChildElements.Count != 3) {
                throw Refuse(context.Sheet.Name, "a data bar that is not two value objects and one color");
            }
            uint minimum = dataBar.MinLength?.Value ?? 10U;
            uint maximum = dataBar.MaxLength?.Value ?? 90U;
            if (minimum > maximum || maximum > 100U) throw Refuse(context.Sheet.Name, $"data bar lengths {minimum} to {maximum}");
            result.Add(new XlsbGeneratedRecord(BrtBeginDatabar, new[] {
                (byte)minimum, (byte)maximum, dataBar.ShowValue?.Value == false ? (byte)0 : (byte)1
            }));
            foreach (ConditionalFormatValueObject value in values) result.Add(CreateValueObject(value, iconSet: false, context));
            result.Add(CreateColor(colors[0], context.Sheet.Name));
            result.Add(new XlsbGeneratedRecord(BrtEndDatabar, Array.Empty<byte>()));
        }

        private static void AppendIconSet(List<XlsbGeneratedRecord> result, IconSet iconSet, XlsbWorksheetFeatureContext context) {
            ConditionalFormatValueObject[] values = iconSet.Elements<ConditionalFormatValueObject>().ToArray();
            IconSetValues set = iconSet.IconSetValue?.Value ?? IconSetValues.ThreeTrafficLights1;
            int kpi = Array.IndexOf(IconSets, set);
            int expected = kpi < 0 ? -1 : kpi < 8 ? 3 : kpi < 13 ? 4 : 5;
            if (XlsbOpenXmlAttributeValidator.FindUnsupported(iconSet, new[] { "iconSet", "showValue", "reverse", "percent" }) != null
                || iconSet.Percent?.Value == false || values.Length != expected || iconSet.ChildElements.Count != values.Length) {
                throw Refuse(context.Sheet.Name, "an icon set that does not match a built-in icon set and its value objects");
            }
            ushort flags = 0;
            if (iconSet.ShowValue?.Value == false) flags |= 0x0002; // fIcon: only the icon is shown.
            if (iconSet.Reverse?.Value == true) flags |= 0x0004;    // fReverse
            using (var payload = new MemoryStream(6)) {
                XlsbPayloadWriter.WriteUInt32(payload, (uint)kpi);
                XlsbPayloadWriter.WriteUInt16(payload, flags);
                result.Add(new XlsbGeneratedRecord(BrtBeginIconSet, payload.ToArray()));
            }
            foreach (ConditionalFormatValueObject value in values) result.Add(CreateValueObject(value, iconSet: true, context));
            result.Add(new XlsbGeneratedRecord(BrtEndIconSet, Array.Empty<byte>()));
        }

        private static XlsbGeneratedRecord CreateValueObject(ConditionalFormatValueObject value, bool iconSet, XlsbWorksheetFeatureContext context) {
            string sheetName = context.Sheet.Name;
            if (value.HasChildren || XlsbOpenXmlAttributeValidator.FindUnsupported(value, iconSet ? new[] { "type", "val", "gte" } : new[] { "type", "val" }) != null) {
                throw Refuse(sheetName, "a conditional formatting value object with extensions");
            }
            ConditionalFormatValueObjectValues kind = value.Type?.Value ?? throw Refuse(sheetName, "a value object without a type");
            int index = Array.FindIndex(ValueObjectTypes, item => item.Type == kind);
            if (index < 0) throw Refuse(sheetName, $"the value object type '{value.Type!.InnerText}'");
            uint type = ValueObjectTypes[index].Value;
            double number = 0D;
            byte[]? formula = null;
            string? text = value.Val?.Value;
            if (type == 7) {
                if (string.IsNullOrEmpty(text)) throw Refuse(sheetName, "a formula value object without a formula");
                formula = EncodeValueObjectFormula(context, text!);
            } else if (type is 1 or 4 or 5) {
                if (string.IsNullOrEmpty(text)) throw Refuse(sheetName, $"a '{value.Type!.InnerText}' value object without a value");
                if (!OfficeIMO.Core.Internal.OfficeRoundTripText.TryParseFinite(text, NumberStyles.Float, CultureInfo.InvariantCulture, out number)) {
                    formula = EncodeValueObjectFormula(context, text!);
                } else if (type != 1 && (number < 0D || number > 100D)) {
                    throw Refuse(sheetName, $"a percent value object of {text}; XLSB allows 0 to 100");
                }
            }
            using var payload = new MemoryStream();
            XlsbPayloadWriter.WriteUInt32(payload, type);
            XlsbPayloadWriter.WriteDouble(payload, number);
            // fSaveGTE and fGTE only apply to icon sets; other value objects store zero, as Excel does.
            XlsbPayloadWriter.WriteUInt32(payload, iconSet ? 1U : 0U);
            XlsbPayloadWriter.WriteUInt32(payload, iconSet && value.GreaterThanOrEqual?.Value != false ? 1U : 0U);
            if (formula == null) {
                XlsbPayloadWriter.WriteUInt32(payload, 0U);
            } else {
                XlsbPayloadWriter.WriteUInt32(payload, BitConverter.ToUInt32(formula, 0)); // cbFmla: the rgce length
                payload.Write(formula, 0, formula.Length);
            }
            return new XlsbGeneratedRecord(BrtCFVO, payload.ToArray());
        }

        // CFParsedFormula keeps relative references as PtgRefN/PtgAreaN offsets from the top-left cell of the first range.
        private static byte[] EncodeRuleFormula(XlsbWorksheetFeatureContext context, string formula, (int Row, int Column) anchor) {
            byte[] encoded = context.EncodeFormula(formula, "conditional formatting");
            byte[] tokens = Tokens(encoded);
            if (HasArrayConstant(tokens)) {
                throw Refuse(context.Sheet.Name, $"the formula '{formula}', because CFParsedFormula cannot hold array constants");
            }
            if (Biff12.XlsbRelativeReferenceTokens.HasSheetReferences(tokens)) {
                throw Refuse(context.Sheet.Name, $"the formula '{formula}', which refers to another sheet (an Office 2010 extension in XLSB)");
            }
            Biff12.XlsbRelativeReferenceTokens.ToOffsets(tokens, anchor.Row, anchor.Column);
            Buffer.BlockCopy(tokens, 0, encoded, 4, tokens.Length);
            return encoded;
        }

        // CFVOParsedFormula (MS-XLSB 2.5.98.7) holds no relative or other-sheet references.
        private static byte[] EncodeValueObjectFormula(XlsbWorksheetFeatureContext context, string formula) {
            byte[] encoded = context.EncodeFormula(formula, "conditional formatting value object");
            byte[] tokens = Tokens(encoded);
            if (HasArrayConstant(tokens)) {
                throw Refuse(context.Sheet.Name, $"the value object formula '{formula}', because CFVOParsedFormula cannot hold array constants");
            }
            if (Biff12.XlsbRelativeReferenceTokens.HasRelativeReferences(tokens) || Biff12.XlsbRelativeReferenceTokens.HasSheetReferences(tokens)) {
                throw Refuse(context.Sheet.Name, $"the value object formula '{formula}', which must use absolute references on the same sheet");
            }
            return encoded;
        }

        // PtgArray (MS-XLSB 2.5.98.23) is 0x20, 0x40 or 0x60; CFParsedFormula and CFVOParsedFormula prohibit it.
        private static bool HasArrayConstant(byte[] tokens) {
            bool found = false;
            Biff12.XlsbFormulaTokenScanner.Visit(tokens, 0, tokens.Length, (offset, token, operand) => {
                if (token is 0x20 or 0x40 or 0x60) found = true;
            });
            return found;
        }

        private static byte[] Tokens(byte[] encoded) {
            int length = checked((int)BitConverter.ToUInt32(encoded, 0));
            var tokens = new byte[length];
            Buffer.BlockCopy(encoded, 4, tokens, 0, length);
            return tokens;
        }

        internal static (int Row, int Column) FindAnchor(string sheetName, string? references) {
            string first = (references ?? string.Empty).Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? string.Empty;
            if (!XlsbPayloadWriter.TryParseRange(first, out int row, out int column, out _, out _)) {
                throw Refuse(sheetName, $"the range '{references}'");
            }
            return (row - 1, column - 1);
        }

        private static XlsbGeneratedRecord CreateColor(DocumentFormat.OpenXml.Spreadsheet.Color color, string sheetName) {
            if (color.HasChildren || XlsbOpenXmlAttributeValidator.FindUnsupported(color, new[] { "auto", "indexed", "rgb", "theme", "tint" }) != null) {
                throw Refuse(sheetName, "a conditional formatting color with extensions");
            }
            using var payload = new MemoryStream(8);
            XlsbStylesheetPartWriter.WriteColor(payload, color);
            return new XlsbGeneratedRecord(BrtColor, payload.ToArray());
        }

        private static NotSupportedException Refuse(string sheetName, string description) =>
            new NotSupportedException($"Native XLSB generation cannot encode conditional formatting on worksheet '{sheetName}': {description}.");
    }
}
