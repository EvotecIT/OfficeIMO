using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using Xm = DocumentFormat.OpenXml.Office.Excel;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Adds a rule from the format-neutral conditional-formatting model.
        /// </summary>
        /// <remarks>
        /// Use <see cref="ExcelConditionalFormattingInfo.Source"/> to select the standard
        /// SpreadsheetML or Office 2010 extension representation. Convenience authoring
        /// methods remain available for common rule families.
        /// </remarks>
        public ExcelConditionalFormattingInfo AddConditionalFormattingRule(
            ExcelConditionalFormattingInfo definition) {
            if (definition == null) throw new ArgumentNullException(nameof(definition));
            string range = ValidateConditionalFormattingRange(definition.Range);
            ValidateConditionalFormattingDefinition(definition, validateFormulas: true, validateVisual: true, validateStyle: true, allowUnknownType: false, updatingExisting: false);
            using var preserveDirectDataSet = _excelDocument.PreserveDirectDataSetSaveCandidateDuringDirtyMarks();
            ExcelConditionalFormattingInfo? result = null;
            WriteLockWorksheetPreparationOnly(() => {
                _excelDocument.EnsureWorkbookThemeAndStyles();
                OpenXmlElement rule;
                OpenXmlElement container;
                if (definition.Source == ExcelConditionalFormattingSource.Office2010Extension) {
                    X14.ConditionalFormattingRule extensionRule = CreateOffice2010ConditionalFormattingRule(definition);
                    var extensionContainer = new X14.ConditionalFormatting(
                        extensionRule,
                        new Xm.ReferenceSequence(range));
                    GetOrCreateOffice2010ConditionalFormattings().Append(extensionContainer);
                    rule = extensionRule;
                    container = extensionContainer;
                } else {
                    ConditionalFormattingRule standardRule = CreateStandardConditionalFormattingRule(definition);
                    var standardContainer = new ConditionalFormatting {
                        SequenceOfReferences = new ListValue<StringValue> { InnerText = range }
                    };
                    standardContainer.Append(standardRule);
                    InsertConditionalFormatting(standardContainer);
                    rule = standardRule;
                    container = standardContainer;
                }

                WorksheetRoot.Save();
                _nextConditionalFormattingPriority = 0;
                result = ReadConditionalFormattingInfo(
                    rule,
                    range,
                    _excelDocument.WorkbookPartRoot.WorkbookStylesPart?.Stylesheet,
                    _excelDocument.WorkbookPartRoot);
                result.OwnerContainer = container;
            });
            return result!;
        }

        /// <summary>
        /// Applies the editable fields of a rule snapshot back to its attached worksheet rule.
        /// Unknown attributes and child elements are retained.
        /// </summary>
        public void UpdateConditionalFormattingRule(ExcelConditionalFormattingInfo rule) {
            if (rule == null) throw new ArgumentNullException(nameof(rule));
            string range = ValidateConditionalFormattingRange(rule.Range);
            bool replaceFormulas = !string.Equals(
                rule.ProjectedFormulaSignature,
                CreateConditionalFormattingFormulaSignature(rule),
                StringComparison.Ordinal) || !string.Equals(rule.ProjectedType, rule.Type, StringComparison.OrdinalIgnoreCase);
            bool replaceVisual = !string.Equals(
                rule.ProjectedVisualSignature,
                CreateConditionalFormattingVisualSignature(rule),
                StringComparison.Ordinal);
            bool replaceStyle = !string.Equals(
                rule.ProjectedStyleSignature,
                CreateConditionalFormattingStyleSignature(rule),
                StringComparison.Ordinal);
            bool unchangedImportedUnknownType = !IsKnownConditionalFormattingType(rule.Type)
                && string.Equals(rule.ProjectedType, rule.Type, StringComparison.Ordinal);
            ValidateConditionalFormattingDefinition(rule, replaceFormulas, replaceVisual, replaceStyle, allowUnknownType: unchangedImportedUnknownType, updatingExisting: true);
            using var preserveDirectDataSet = _excelDocument.PreserveDirectDataSetSaveCandidateDuringDirtyMarks();
            WriteLock(() => {
                OpenXmlElement backing = RequireAttachedConditionalFormattingRule(rule);
                if (backing is ConditionalFormattingRule standardRule) {
                    if (rule.OwnerContainer is not ConditionalFormatting standardContainer) {
                        throw new InvalidOperationException("The standard conditional-formatting rule owner is no longer attached.");
                    }
                    standardContainer = SplitStandardConditionalFormattingOwnerForRange(
                        standardContainer,
                        standardRule,
                        range);
                    rule.OwnerContainer = standardContainer;
                    standardContainer.SequenceOfReferences = new ListValue<StringValue> { InnerText = range };
                    ApplyStandardConditionalFormattingRule(standardRule, rule, replaceFormulas: replaceFormulas, replaceVisual: replaceVisual, replaceStyle: replaceStyle);
                } else if (backing is X14.ConditionalFormattingRule extensionRule) {
                    if (rule.OwnerContainer is not X14.ConditionalFormatting extensionContainer) {
                        throw new InvalidOperationException("The Office extension conditional-formatting rule owner is no longer attached.");
                    }
                    extensionContainer = SplitOffice2010ConditionalFormattingOwnerForRange(
                        extensionContainer,
                        extensionRule,
                        range);
                    rule.OwnerContainer = extensionContainer;
                    Xm.ReferenceSequence target = extensionContainer.GetFirstChild<Xm.ReferenceSequence>()
                        ?? extensionContainer.AppendChild(new Xm.ReferenceSequence());
                    target.Text = range;
                    ApplyOffice2010ConditionalFormattingRule(extensionRule, rule, replaceFormulas: replaceFormulas, replaceVisual: replaceVisual, replaceStyle: replaceStyle);
                } else {
                    throw new InvalidOperationException("The conditional-formatting rule type is unsupported.");
                }

                _nextConditionalFormattingPriority = 0;
                WorksheetRoot.Save();
                CaptureConditionalFormattingProjectionSignatures(rule);
            });
        }

        /// <summary>
        /// Clones a rule, including unrecognized attributes and child markup, to another range.
        /// </summary>
        public ExcelConditionalFormattingInfo CloneConditionalFormattingRule(
            ExcelConditionalFormattingInfo rule,
            string targetRange,
            int? priority = null) {
            if (rule == null) throw new ArgumentNullException(nameof(rule));
            string range = ValidateConditionalFormattingRange(targetRange);
            using var preserveDirectDataSet = _excelDocument.PreserveDirectDataSetSaveCandidateDuringDirtyMarks();
            ExcelConditionalFormattingInfo? result = null;
            WriteLock(() => {
                OpenXmlElement backing = RequireAttachedConditionalFormattingRule(rule);
                OpenXmlElement clone = backing.CloneNode(true);
                int assignedPriority = priority ?? GetNextConditionalFormattingPriority();
                SetConditionalFormattingPriority(clone, assignedPriority);
                OpenXmlElement container;
                if (clone is ConditionalFormattingRule standardClone) {
                    ConditionalFormatting standardContainer = backing.Parent is ConditionalFormatting owner
                        ? (ConditionalFormatting)owner.CloneNode(true)
                        : new ConditionalFormatting();
                    foreach (ConditionalFormattingRule ownerRule in
                        standardContainer.Elements<ConditionalFormattingRule>().ToList()) ownerRule.Remove();
                    standardContainer.SequenceOfReferences = new ListValue<StringValue> { InnerText = range };
                    InsertBeforeRuleExtension(standardContainer, standardClone);
                    InsertConditionalFormatting(standardContainer);
                    container = standardContainer;
                } else if (clone is X14.ConditionalFormattingRule extensionClone) {
                    extensionClone.Id = "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}";
                    X14.ConditionalFormatting extensionContainer = backing.Parent is X14.ConditionalFormatting owner
                        ? (X14.ConditionalFormatting)owner.CloneNode(true)
                        : new X14.ConditionalFormatting();
                    foreach (X14.ConditionalFormattingRule ownerRule in
                        extensionContainer.Elements<X14.ConditionalFormattingRule>().ToList()) ownerRule.Remove();
                    Xm.ReferenceSequence target = extensionContainer.GetFirstChild<Xm.ReferenceSequence>()
                        ?? extensionContainer.AppendChild(new Xm.ReferenceSequence());
                    target.Text = range;
                    extensionContainer.InsertBefore(extensionClone, target);
                    GetOrCreateOffice2010ConditionalFormattings().Append(extensionContainer);
                    container = extensionContainer;
                } else {
                    throw new InvalidOperationException("The conditional-formatting rule type is unsupported.");
                }

                WorksheetRoot.Save();
                result = ReadConditionalFormattingInfo(
                    clone,
                    range,
                    _excelDocument.WorkbookPartRoot.WorkbookStylesPart?.Stylesheet,
                    _excelDocument.WorkbookPartRoot);
                result.OwnerContainer = container;
            });
            return result!;
        }

        /// <summary>Removes one attached conditional-formatting rule.</summary>
        public void RemoveConditionalFormattingRule(ExcelConditionalFormattingInfo rule) {
            if (rule == null) throw new ArgumentNullException(nameof(rule));
            WriteLock(() => {
                OpenXmlElement backing = RequireAttachedConditionalFormattingRule(rule);
                OpenXmlElement? container = backing.Parent;
                backing.Remove();
                if (container != null && !HasConditionalFormattingRuleChildren(container)) container.Remove();
                CleanupEmptyConditionalFormattingOwners();
                rule.BackingRule = null;
                rule.OwnerContainer = null;
                _nextConditionalFormattingPriority = 0;
                WorksheetRoot.Save();
            });
        }

        /// <summary>
        /// Reassigns consecutive priorities in the supplied order. The sequence must contain
        /// every currently attached rule exactly once.
        /// </summary>
        public void ReorderConditionalFormattingRules(
            IReadOnlyList<ExcelConditionalFormattingInfo> orderedRules) {
            if (orderedRules == null) throw new ArgumentNullException(nameof(orderedRules));
            WriteLock(() => {
                List<OpenXmlElement> current = EnumerateAttachedConditionalFormattingRules().ToList();
                if (current.Count != orderedRules.Count) {
                    throw new ArgumentException("The ordered rule list must contain every worksheet conditional-formatting rule exactly once.", nameof(orderedRules));
                }

                var supplied = new HashSet<OpenXmlElement>();
                var orderedBacking = new List<OpenXmlElement>(orderedRules.Count);
                for (int index = 0; index < orderedRules.Count; index++) {
                    OpenXmlElement backing = RequireAttachedConditionalFormattingRule(orderedRules[index]);
                    if (!supplied.Add(backing)) {
                        throw new ArgumentException("The ordered rule list contains a duplicate rule.", nameof(orderedRules));
                    }
                    orderedBacking.Add(backing);
                }

                if (supplied.Count != current.Count || current.Any(candidate => !supplied.Contains(candidate))) {
                    throw new ArgumentException("The ordered rule list contains a rule from another worksheet.", nameof(orderedRules));
                }

                for (int index = 0; index < orderedBacking.Count; index++) {
                    SetConditionalFormattingPriority(orderedBacking[index], index + 1);
                    orderedRules[index].Priority = index + 1;
                }

                _nextConditionalFormattingPriority = orderedRules.Count + 1;
                WorksheetRoot.Save();
            });
        }

        private ConditionalFormattingRule CreateStandardConditionalFormattingRule(
            ExcelConditionalFormattingInfo definition) {
            var rule = new ConditionalFormattingRule();
            ApplyStandardConditionalFormattingRule(rule, definition, creating: true);
            return rule;
        }

        private X14.ConditionalFormattingRule CreateOffice2010ConditionalFormattingRule(
            ExcelConditionalFormattingInfo definition) {
            var rule = new X14.ConditionalFormattingRule {
                Id = "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}"
            };
            ApplyOffice2010ConditionalFormattingRule(rule, definition, creating: true);
            return rule;
        }

        private void ApplyStandardConditionalFormattingRule(
            ConditionalFormattingRule rule,
            ExcelConditionalFormattingInfo definition,
            bool creating = false,
            bool replaceFormulas = true,
            bool replaceVisual = true,
            bool replaceStyle = true) {
            SetRuleAttribute(rule, "type", ToConditionalFormattingToken(definition.Type));
            SetRuleAttribute(rule, "operator", ToConditionalFormattingToken(definition.Operator));
            SetRuleAttribute(rule, "text", definition.Text);
            SetRuleAttribute(rule, "timePeriod", ToConditionalFormattingToken(definition.TimePeriod));
            rule.Priority = definition.Priority > 0 ? definition.Priority : creating ? GetNextConditionalFormattingPriority() : 1;
            rule.StopIfTrue = definition.StopIfTrue;
            uint? differentialFormatId = replaceStyle
                ? CreateOrReuseStandardConditionalDifferentialFormat(definition)
                : definition.DifferentialFormatId;
            rule.FormatId = differentialFormatId;
            definition.DifferentialFormatId = differentialFormatId;
            SetRuleAttribute(rule, "rank", definition.TopBottomRank?.ToString(CultureInfo.InvariantCulture));
            SetRuleAttribute(rule, "bottom", definition.TopBottomBottom ? "1" : null);
            SetRuleAttribute(rule, "percent", definition.TopBottomPercent ? "1" : null);
            SetRuleAttribute(rule, "aboveAverage", definition.AboveAverageAbove ? null : "0");
            SetRuleAttribute(rule, "equalAverage", definition.AboveAverageEqual ? "1" : null);
            SetRuleAttribute(rule, "stdDev", definition.AboveAverageStdDev?.ToString(CultureInfo.InvariantCulture));
            if (replaceFormulas) ReplaceRuleFormulas<Formula>(rule, definition.Formulas, value => new Formula(value));
            if (replaceVisual) ReplaceStandardConditionalFormattingVisual(rule, definition);
        }

        private void ApplyOffice2010ConditionalFormattingRule(
            X14.ConditionalFormattingRule rule,
            ExcelConditionalFormattingInfo definition,
            bool creating = false,
            bool replaceFormulas = true,
            bool replaceVisual = true,
            bool replaceStyle = true) {
            SetRuleAttribute(rule, "type", ToConditionalFormattingToken(definition.Type));
            SetRuleAttribute(rule, "operator", ToConditionalFormattingToken(definition.Operator));
            SetRuleAttribute(rule, "text", definition.Text);
            SetRuleAttribute(rule, "timePeriod", ToConditionalFormattingToken(definition.TimePeriod));
            rule.Priority = definition.Priority > 0 ? definition.Priority : creating ? GetNextConditionalFormattingPriority() : 1;
            rule.StopIfTrue = definition.StopIfTrue;
            SetRuleAttribute(rule, "rank", definition.TopBottomRank?.ToString(CultureInfo.InvariantCulture));
            SetRuleAttribute(rule, "bottom", definition.TopBottomBottom ? "1" : null);
            SetRuleAttribute(rule, "percent", definition.TopBottomPercent ? "1" : null);
            SetRuleAttribute(rule, "aboveAverage", definition.AboveAverageAbove ? null : "0");
            SetRuleAttribute(rule, "equalAverage", definition.AboveAverageEqual ? "1" : null);
            SetRuleAttribute(rule, "stdDev", definition.AboveAverageStdDev?.ToString(CultureInfo.InvariantCulture));
            if (replaceFormulas) ReplaceRuleFormulas<Xm.Formula>(rule, definition.Formulas, value => new Xm.Formula(value));
            if (replaceVisual) ReplaceOffice2010ConditionalFormattingVisual(rule, definition);
            if (replaceStyle) ApplyOffice2010ConditionalDifferentialFormat(rule, definition);
        }

        private static void ReplaceRuleFormulas<TFormula>(
            OpenXmlCompositeElement rule,
            IReadOnlyList<string> formulas,
            Func<string, TFormula> create)
            where TFormula : OpenXmlElement {
            foreach (TFormula formula in rule.Elements<TFormula>().ToList()) formula.Remove();
            OpenXmlElement? before = rule.ChildElements.FirstOrDefault();
            foreach (string formula in formulas ?? Array.Empty<string>()) {
                if (string.IsNullOrWhiteSpace(formula)) {
                    throw new ArgumentException("Conditional-formatting formulas cannot be empty.", nameof(formulas));
                }
                TFormula created = create(formula);
                if (before == null) rule.Append(created);
                else rule.InsertBefore(created, before);
            }
        }

        private static void ReplaceStandardConditionalFormattingVisual(
            ConditionalFormattingRule rule,
            ExcelConditionalFormattingInfo definition) {
            if (string.Equals(definition.Type, "ColorScale", StringComparison.OrdinalIgnoreCase)) {
                ColorScale? scale = rule.GetFirstChild<ColorScale>();
                RemoveOtherConditionalFormattingVisuals(rule, scale);
                if (scale == null) {
                    scale = new ColorScale();
                    InsertBeforeRuleExtension(rule, scale);
                }
                SynchronizeStandardThresholds(scale, definition.ColorScaleThresholds, iconThresholds: null);
                SynchronizeStandardColors(scale, definition.ColorScaleColors);
            } else if (string.Equals(definition.Type, "DataBar", StringComparison.OrdinalIgnoreCase)) {
                DataBar? dataBar = rule.GetFirstChild<DataBar>();
                RemoveOtherConditionalFormattingVisuals(rule, dataBar);
                if (dataBar == null) {
                    dataBar = new DataBar();
                    InsertBeforeRuleExtension(rule, dataBar);
                }
                dataBar.ShowValue = definition.DataBarShowValue;
                SynchronizeStandardThresholds(dataBar, definition.DataBarThresholds, iconThresholds: null);
                if (!string.IsNullOrWhiteSpace(definition.DataBarColor)) {
                    DocumentFormat.OpenXml.Spreadsheet.Color color = dataBar.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Color>()
                        ?? dataBar.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Color());
                    SetStandardConditionalFormattingColor(color, definition.DataBarColor!);
                }
            } else if (string.Equals(definition.Type, "IconSet", StringComparison.OrdinalIgnoreCase)) {
                IconSet? iconSet = rule.GetFirstChild<IconSet>();
                RemoveOtherConditionalFormattingVisuals(rule, iconSet);
                if (iconSet == null) {
                    iconSet = new IconSet();
                    InsertBeforeRuleExtension(rule, iconSet);
                }
                iconSet.ShowValue = definition.IconSetShowValue;
                iconSet.Reverse = definition.IconSetReverse;
                SetRuleAttribute(iconSet, "iconSet", ToConditionalIconSetToken(definition.IconSet));
                SynchronizeStandardThresholds(iconSet, thresholds: null, definition.IconSetThresholds);
            } else {
                RemoveOtherConditionalFormattingVisuals(rule, retained: null);
            }
        }

        private static void ReplaceOffice2010ConditionalFormattingVisual(
            X14.ConditionalFormattingRule rule,
            ExcelConditionalFormattingInfo definition) {
            if (string.Equals(definition.Type, "ColorScale", StringComparison.OrdinalIgnoreCase)) {
                X14.ColorScale? scale = rule.GetFirstChild<X14.ColorScale>();
                RemoveOtherOffice2010ConditionalFormattingVisuals(rule, scale);
                if (scale == null) {
                    scale = new X14.ColorScale();
                    InsertBeforeRuleExtension(rule, scale);
                }
                SynchronizeOffice2010Thresholds(scale, definition.ColorScaleThresholds, iconThresholds: null);
                SynchronizeOffice2010Colors(scale, definition.ColorScaleColors);
            } else if (string.Equals(definition.Type, "DataBar", StringComparison.OrdinalIgnoreCase)) {
                X14.DataBar? dataBar = rule.GetFirstChild<X14.DataBar>();
                RemoveOtherOffice2010ConditionalFormattingVisuals(rule, dataBar);
                if (dataBar == null) {
                    dataBar = new X14.DataBar();
                    InsertBeforeRuleExtension(rule, dataBar);
                }
                dataBar.ShowValue = definition.DataBarShowValue;
                dataBar.MinLength = definition.DataBarMinimumLength;
                dataBar.MaxLength = definition.DataBarMaximumLength;
                dataBar.Border = definition.DataBarBorder;
                dataBar.Gradient = definition.DataBarGradient;
                dataBar.NegativeBarColorSameAsPositive = definition.DataBarNegativeColorSameAsPositive;
                dataBar.NegativeBarBorderColorSameAsPositive = definition.DataBarNegativeBorderColorSameAsPositive;
                SetRuleAttribute(dataBar, "direction", ToConditionalFormattingToken(definition.DataBarDirection));
                SetRuleAttribute(dataBar, "axisPosition", ToConditionalFormattingToken(definition.DataBarAxisPosition));
                SynchronizeOffice2010Thresholds(dataBar, definition.DataBarThresholds, iconThresholds: null);
                SetOffice2010DataBarColor<X14.FillColor>(dataBar, definition.DataBarColor);
                SetOffice2010DataBarColor<X14.BorderColor>(dataBar, definition.DataBarBorderColor);
                SetOffice2010DataBarColor<X14.NegativeFillColor>(dataBar, definition.DataBarNegativeColor);
                SetOffice2010DataBarColor<X14.NegativeBorderColor>(dataBar, definition.DataBarNegativeBorderColor);
                SetOffice2010DataBarColor<X14.BarAxisColor>(dataBar, definition.DataBarAxisColor);
            } else if (string.Equals(definition.Type, "IconSet", StringComparison.OrdinalIgnoreCase)) {
                X14.IconSet? iconSet = rule.GetFirstChild<X14.IconSet>();
                RemoveOtherOffice2010ConditionalFormattingVisuals(rule, iconSet);
                if (iconSet == null) {
                    iconSet = new X14.IconSet();
                    InsertBeforeRuleExtension(rule, iconSet);
                }
                iconSet.ShowValue = definition.IconSetShowValue;
                iconSet.Reverse = definition.IconSetReverse;
                iconSet.Percent = definition.IconSetPercent;
                iconSet.Custom = definition.IconSetCustom;
                SetRuleAttribute(iconSet, "iconSet", ToConditionalIconSetToken(definition.IconSet));
                SynchronizeOffice2010Thresholds(iconSet, thresholds: null, definition.IconSetThresholds);
                SynchronizeOffice2010CustomIcons(iconSet, definition.CustomIcons);
            } else {
                RemoveOtherOffice2010ConditionalFormattingVisuals(rule, retained: null);
            }
        }

        private static void RemoveOtherConditionalFormattingVisuals(
            ConditionalFormattingRule rule,
            OpenXmlElement? retained) {
            foreach (OpenXmlElement visual in rule.ChildElements
                .Where(element => element is ColorScale || element is DataBar || element is IconSet)
                .Where(element => !ReferenceEquals(element, retained))
                .ToList()) visual.Remove();
        }

        private static void RemoveOtherOffice2010ConditionalFormattingVisuals(
            X14.ConditionalFormattingRule rule,
            OpenXmlElement? retained) {
            foreach (OpenXmlElement visual in rule.ChildElements
                .Where(element => element is X14.ColorScale || element is X14.DataBar || element is X14.IconSet)
                .Where(element => !ReferenceEquals(element, retained))
                .ToList()) visual.Remove();
        }

        private static void SynchronizeStandardThresholds(
            OpenXmlCompositeElement visual,
            IReadOnlyList<ExcelConditionalFormatThreshold>? thresholds,
            IReadOnlyList<ExcelConditionalIconSetThreshold>? iconThresholds) {
            int count = thresholds?.Count ?? iconThresholds?.Count ?? 0;
            List<ConditionalFormatValueObject> existing = visual.Elements<ConditionalFormatValueObject>().ToList();
            if (existing.Count != count) {
                foreach (ConditionalFormatValueObject value in existing) value.Remove();
                existing.Clear();
                OpenXmlElement? before = visual.ChildElements.FirstOrDefault();
                for (int index = 0; index < count; index++) {
                    var created = new ConditionalFormatValueObject();
                    if (before == null) visual.Append(created);
                    else visual.InsertBefore(created, before);
                    existing.Add(created);
                }
            }
            for (int index = 0; index < count; index++) {
                ConditionalFormatValueObject value = existing[index];
                string? type = thresholds != null ? thresholds[index].Type : iconThresholds![index].Type;
                string? text = thresholds != null ? thresholds[index].Value : iconThresholds![index].Value;
                if (string.IsNullOrWhiteSpace(text)) value.RemoveAttribute("val", string.Empty);
                else value.Val = text;
                value.GreaterThanOrEqual = iconThresholds != null ? iconThresholds[index].GreaterThanOrEqual : null;
                SetRuleAttribute(value, "type", ToConditionalFormatValueToken(type));
            }
        }

        private static void SynchronizeOffice2010Thresholds(
            OpenXmlCompositeElement visual,
            IReadOnlyList<ExcelConditionalFormatThreshold>? thresholds,
            IReadOnlyList<ExcelConditionalIconSetThreshold>? iconThresholds) {
            int count = thresholds?.Count ?? iconThresholds?.Count ?? 0;
            List<X14.ConditionalFormattingValueObject> existing = visual.Elements<X14.ConditionalFormattingValueObject>().ToList();
            if (existing.Count != count) {
                foreach (X14.ConditionalFormattingValueObject value in existing) value.Remove();
                existing.Clear();
                OpenXmlElement? before = visual.ChildElements.FirstOrDefault();
                for (int index = 0; index < count; index++) {
                    var created = new X14.ConditionalFormattingValueObject();
                    if (before == null) visual.Append(created);
                    else visual.InsertBefore(created, before);
                    existing.Add(created);
                }
            }
            for (int index = 0; index < count; index++) {
                X14.ConditionalFormattingValueObject value = existing[index];
                string? type = thresholds != null ? thresholds[index].Type : iconThresholds![index].Type;
                string? text = thresholds != null ? thresholds[index].Value : iconThresholds![index].Value;
                value.GreaterThanOrEqual = iconThresholds != null ? iconThresholds[index].GreaterThanOrEqual : null;
                SetRuleAttribute(value, "type", ToConditionalFormatValueToken(type));
                Xm.Formula? formula = value.GetFirstChild<Xm.Formula>();
                if (string.IsNullOrWhiteSpace(text)) formula?.Remove();
                else if (formula == null) value.PrependChild(new Xm.Formula(text!));
                else formula.Text = text!;
            }
        }

        private static void SynchronizeStandardColors(
            ColorScale scale,
            IReadOnlyList<string> colors) {
            if (colors.Count == 0) return;
            List<DocumentFormat.OpenXml.Spreadsheet.Color> existing = scale
                .Elements<DocumentFormat.OpenXml.Spreadsheet.Color>().ToList();
            if (existing.Count != colors.Count) {
                if (colors.Any(string.IsNullOrWhiteSpace)) {
                    throw new ArgumentException("Imported theme, indexed, or automatic color slots can only be preserved in their original positions.", nameof(colors));
                }
                foreach (DocumentFormat.OpenXml.Spreadsheet.Color color in existing) color.Remove();
                existing = colors.Select(_ => scale.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Color())).ToList();
            }
            for (int index = 0; index < colors.Count; index++) {
                if (!string.IsNullOrWhiteSpace(colors[index])) SetStandardConditionalFormattingColor(existing[index], colors[index]);
            }
        }

        private static void SetStandardConditionalFormattingColor(
            DocumentFormat.OpenXml.Spreadsheet.Color color,
            string value) {
            string normalized = NormalizeHexColor(value);
            if (string.Equals(color.Rgb?.Value, normalized, StringComparison.OrdinalIgnoreCase)) return;
            color.Auto = null;
            color.Indexed = null;
            color.Theme = null;
            color.Tint = null;
            color.Rgb = normalized;
        }

        private static void SynchronizeOffice2010Colors(
            X14.ColorScale scale,
            IReadOnlyList<string> colors) {
            if (colors.Count == 0) return;
            List<X14.Color> existing = scale.Elements<X14.Color>().ToList();
            if (existing.Count != colors.Count) {
                if (colors.Any(string.IsNullOrWhiteSpace)) {
                    throw new ArgumentException("Imported theme, indexed, or automatic color slots can only be preserved in their original positions.", nameof(colors));
                }
                foreach (X14.Color color in existing) color.Remove();
                existing = colors.Select(_ => scale.AppendChild(new X14.Color())).ToList();
            }
            for (int index = 0; index < colors.Count; index++) {
                if (!string.IsNullOrWhiteSpace(colors[index])) SetOffice2010ConditionalFormattingColor(existing[index], colors[index]);
            }
        }

        private static void SetOffice2010DataBarColor<TColor>(X14.DataBar dataBar, string? value)
            where TColor : X14.ColorType, new() {
            if (string.IsNullOrWhiteSpace(value)) return;
            TColor color = dataBar.GetFirstChild<TColor>() ?? new TColor();
            SetOffice2010ConditionalFormattingColor(color, value!);
            if (color.Parent == null) InsertOffice2010DataBarColor(dataBar, color);
        }

        private static void SetOffice2010ConditionalFormattingColor(X14.ColorType color, string value) {
            string normalized = NormalizeHexColor(value);
            if (string.Equals(color.Rgb?.Value, normalized, StringComparison.OrdinalIgnoreCase)) return;
            color.Auto = null;
            color.Indexed = null;
            color.Theme = null;
            color.Tint = null;
            color.Rgb = normalized;
        }

        private static void InsertOffice2010DataBarColor(X14.DataBar dataBar, X14.ColorType color) {
            int rank = GetOffice2010DataBarColorRank(color);
            OpenXmlElement? before = dataBar.ChildElements.FirstOrDefault(child =>
                child is X14.ColorType existing && GetOffice2010DataBarColorRank(existing) > rank ||
                child is X14.ExtensionList);
            if (before == null) dataBar.Append(color);
            else dataBar.InsertBefore(color, before);
        }

        private static int GetOffice2010DataBarColorRank(X14.ColorType color) => color switch {
            X14.FillColor => 0,
            X14.BorderColor => 1,
            X14.NegativeFillColor => 2,
            X14.NegativeBorderColor => 3,
            X14.BarAxisColor => 4,
            _ => 5
        };

        private static void SynchronizeOffice2010CustomIcons(
            X14.IconSet iconSet,
            IReadOnlyList<ExcelConditionalFormatIcon> icons) {
            List<X14.ConditionalFormattingIcon> existing = iconSet.Elements<X14.ConditionalFormattingIcon>().ToList();
            if (existing.Count != icons.Count) {
                foreach (X14.ConditionalFormattingIcon icon in existing) icon.Remove();
                existing = icons.Select(_ => iconSet.AppendChild(new X14.ConditionalFormattingIcon())).ToList();
            }
            for (int index = 0; index < icons.Count; index++) {
                existing[index].IconId = icons[index].IconId;
                SetRuleAttribute(existing[index], "iconSet", ToConditionalIconSetToken(icons[index].IconSet));
            }
        }

        private static void InsertBeforeRuleExtension(OpenXmlCompositeElement rule, OpenXmlElement child) {
            OpenXmlElement? extension = rule.ChildElements.FirstOrDefault(element =>
                element is DocumentFormat.OpenXml.Spreadsheet.ExtensionList || element is X14.ExtensionList);
            if (extension == null) rule.Append(child);
            else rule.InsertBefore(child, extension);
        }

        private X14.ConditionalFormattings GetOrCreateOffice2010ConditionalFormattings() {
            WorksheetExtensionList extensions = WorksheetRoot.GetFirstChild<WorksheetExtensionList>()
                ?? WorksheetRoot.AppendChild(new WorksheetExtensionList());
            WorksheetExtension? extension = extensions.Elements<WorksheetExtension>().FirstOrDefault(candidate =>
                string.Equals(candidate.Uri?.Value, Office2010ConditionalFormattingExtensionUri, StringComparison.OrdinalIgnoreCase));
            if (extension == null) {
                extension = new WorksheetExtension { Uri = Office2010ConditionalFormattingExtensionUri };
                extensions.Append(extension);
            }
            X14.ConditionalFormattings? formattings = extension.GetFirstChild<X14.ConditionalFormattings>();
            if (formattings == null) {
                formattings = new X14.ConditionalFormattings();
                extension.Append(formattings);
            }
            return formattings;
        }

        private OpenXmlElement RequireAttachedConditionalFormattingRule(ExcelConditionalFormattingInfo rule) {
            OpenXmlElement? backing = rule.BackingRule;
            OpenXmlElement? owner = rule.OwnerContainer;
            if (!ReferenceEquals(rule.OwnerSheet, this)
                || backing == null
                || owner == null
                || !ReferenceEquals(backing.Parent, owner)) {
                throw new InvalidOperationException("The conditional-formatting rule is detached or belongs to another worksheet.");
            }
            OpenXmlElement? ancestor = owner;
            while (ancestor != null && !ReferenceEquals(ancestor, WorksheetRoot)) ancestor = ancestor.Parent;
            if (!ReferenceEquals(ancestor, WorksheetRoot)) {
                throw new InvalidOperationException("The conditional-formatting rule is detached or belongs to another worksheet.");
            }
            return backing;
        }

        private IEnumerable<OpenXmlElement> EnumerateAttachedConditionalFormattingRules() {
            foreach (ConditionalFormattingRule rule in WorksheetRoot
                .Elements<ConditionalFormatting>()
                .SelectMany(container => container.Elements<ConditionalFormattingRule>())) yield return rule;
            WorksheetExtensionList? extensions = WorksheetRoot.GetFirstChild<WorksheetExtensionList>();
            if (extensions != null) {
                foreach (X14.ConditionalFormattingRule rule in extensions
                    .Descendants<X14.ConditionalFormattingRule>()) yield return rule;
            }
        }

        private static bool HasConditionalFormattingRuleChildren(OpenXmlElement container) =>
            container.Elements<ConditionalFormattingRule>().Any() ||
            container.Elements<X14.ConditionalFormattingRule>().Any();

        private static ConditionalFormatting SplitStandardConditionalFormattingOwnerForRange(
            ConditionalFormatting owner,
            ConditionalFormattingRule rule,
            string range) {
            string currentRange = owner.SequenceOfReferences?.InnerText ?? string.Empty;
            if (string.Equals(currentRange, range, StringComparison.Ordinal) ||
                owner.Elements<ConditionalFormattingRule>().Count() <= 1) return owner;

            var split = (ConditionalFormatting)owner.CloneNode(true);
            foreach (ConditionalFormattingRule clonedRule in split.Elements<ConditionalFormattingRule>().ToList()) clonedRule.Remove();
            split.SequenceOfReferences = new ListValue<StringValue> { InnerText = range };
            rule.Remove();
            InsertBeforeRuleExtension(split, rule);
            owner.Parent!.InsertAfter(split, owner);
            return split;
        }

        private static X14.ConditionalFormatting SplitOffice2010ConditionalFormattingOwnerForRange(
            X14.ConditionalFormatting owner,
            X14.ConditionalFormattingRule rule,
            string range) {
            string currentRange = owner.GetFirstChild<Xm.ReferenceSequence>()?.Text ?? string.Empty;
            if (string.Equals(currentRange, range, StringComparison.Ordinal) ||
                owner.Elements<X14.ConditionalFormattingRule>().Count() <= 1) return owner;

            var split = (X14.ConditionalFormatting)owner.CloneNode(true);
            foreach (X14.ConditionalFormattingRule clonedRule in split.Elements<X14.ConditionalFormattingRule>().ToList()) clonedRule.Remove();
            Xm.ReferenceSequence target = split.GetFirstChild<Xm.ReferenceSequence>()
                ?? split.AppendChild(new Xm.ReferenceSequence());
            target.Text = range;
            rule.Remove();
            split.InsertBefore(rule, target);
            owner.Parent!.InsertAfter(split, owner);
            return split;
        }

        private void CleanupEmptyConditionalFormattingOwners() {
            foreach (X14.ConditionalFormattings formattings in WorksheetRoot
                .Descendants<X14.ConditionalFormattings>().ToList()) {
                if (!formattings.Elements<X14.ConditionalFormatting>().Any()) formattings.Remove();
            }
            CleanupEmptyMetadataExtensions();
        }

        private static void SetConditionalFormattingPriority(OpenXmlElement rule, int priority) {
            if (priority <= 0) throw new ArgumentOutOfRangeException(nameof(priority));
            if (rule is ConditionalFormattingRule standard) standard.Priority = priority;
            else if (rule is X14.ConditionalFormattingRule extension) extension.Priority = priority;
            else throw new InvalidOperationException("The conditional-formatting rule type is unsupported.");
        }

        private static string ValidateConditionalFormattingRange(string range) {
            if (string.IsNullOrWhiteSpace(range)) throw new ArgumentException("A conditional-formatting range is required.", nameof(range));
            var normalized = new List<string>();
            foreach (ReferenceListPart part in SplitReferenceList(range.Trim())) {
                if (!TryParseReference(part, out _)) {
                    throw new ArgumentException($"Invalid conditional-formatting A1 reference '{part}'.", nameof(range));
                }
                normalized.Add(part.ToString());
            }
            if (normalized.Count == 0) throw new ArgumentException("A conditional-formatting range is required.", nameof(range));
            return string.Join(" ", normalized);
        }

        private static void SetRuleAttribute(OpenXmlElement element, string name, string? value) {
            if (string.IsNullOrWhiteSpace(value)) element.RemoveAttribute(name, string.Empty);
            else element.SetAttribute(new OpenXmlAttribute(name, string.Empty, value));
        }

        private static string? ToConditionalFormattingToken(string? value) {
            if (string.IsNullOrWhiteSpace(value)) return null;
            return char.ToLowerInvariant(value![0]) + value.Substring(1);
        }

        private static string? ToConditionalFormatValueToken(string? value) {
            if (string.IsNullOrWhiteSpace(value)) return null;
            if (string.Equals(value, "Minimum", StringComparison.OrdinalIgnoreCase)) return "min";
            if (string.Equals(value, "Maximum", StringComparison.OrdinalIgnoreCase)) return "max";
            if (string.Equals(value, "Number", StringComparison.OrdinalIgnoreCase)) return "num";
            return ToConditionalFormattingToken(value);
        }

        private static string? ToConditionalIconSetToken(string? value) {
            string? normalized = NormalizeConditionalIconSetName(value);
            if (string.IsNullOrWhiteSpace(normalized)) return null;
            if (normalized!.StartsWith("Three", StringComparison.Ordinal)) return "3" + normalized.Substring(5);
            if (normalized.StartsWith("Four", StringComparison.Ordinal)) return "4" + normalized.Substring(4);
            if (normalized.StartsWith("Five", StringComparison.Ordinal)) return "5" + normalized.Substring(4);
            return ToConditionalFormattingToken(normalized);
        }

    }
}
