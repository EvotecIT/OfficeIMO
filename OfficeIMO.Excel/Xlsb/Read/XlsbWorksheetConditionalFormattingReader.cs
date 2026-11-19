using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Projection;
using OfficeIMO.Excel.Xlsb.Styles;
using W = OfficeIMO.Excel.Xlsb.Write.XlsbWorksheetConditionalFormattingWriter;

namespace OfficeIMO.Excel.Xlsb.Read {
    /// <summary>
    /// Decodes CONDITIONALFORMATTING collections (MS-XLSB 2.4.34, 2.4.23) into conditionalFormatting elements. A collection
    /// whose content has no SpreadsheetML equivalent here, such as an Office 2010 rule extension, is left unprojected and
    /// recorded in <see cref="Model.XlsbSheet.HasUnprojectedConditionalFormatting"/> so a rewrite keeps or refuses it.
    /// </summary>
    internal sealed class XlsbWorksheetConditionalFormattingReader {
        private const int BrtFrtBegin = 35;
        private const int BrtFrtEnd = 36;
        private const int BrtBeginConditionalFormattings = 1135;

        private readonly XlsbWorksheetFeatureReader _owner;
        private ConditionalFormatting? _collection;
        private int _declaredRules;
        private ConditionalFormattingRule? _rule;
        private OpenXmlCompositeElement? _visual;
        private int _visualEnd;
        private readonly List<ConditionalFormatValueObject> _values = new List<ConditionalFormatValueObject>();
        private readonly List<DocumentFormat.OpenXml.Spreadsheet.Color> _colors = new List<DocumentFormat.OpenXml.Spreadsheet.Color>();
        private bool _unprojected;
        private int _futureDepth;
        private XlsbRecord? _unprojectedRecord;

        internal XlsbWorksheetConditionalFormattingReader(XlsbWorksheetFeatureReader owner) => _owner = owner;

        internal bool TryRead(XlsbRecord record) {
            if (_collection == null) {
                if (record.Type == BrtBeginConditionalFormattings) {
                    // Office 2010 rules share priorities and rule links with these collections; keep them as one unit.
                    _owner.Sheet.HasUnprojectedConditionalFormatting = true;
                    return false;
                }
                if (record.Type != W.BrtBeginConditionalFormatting) return false;
                BeginCollection(record);
                return true;
            }
            if (_futureDepth > 0) {
                if (record.Type == BrtFrtBegin) _futureDepth++;
                else if (record.Type == BrtFrtEnd) _futureDepth--;
                return true;
            }
            switch (record.Type) {
                case W.BrtBeginCFRule:
                    if (_rule != null) throw Invalid("a nested BrtBeginCFRule record");
                    _rule = ReadRule(record);
                    return true;
                case W.BrtEndCFRule:
                    if (_rule == null || _visual != null) throw Invalid("an unmatched BrtEndCFRule record");
                    if (IsVisual(_rule) && !_rule.HasChildren) MarkUnprojected(record);
                    _collection.Append(_rule);
                    _rule = null;
                    return true;
                case W.BrtBeginColorScale:
                case W.BrtBeginDatabar:
                case W.BrtBeginIconSet:
                    if (_rule == null || _visual != null || _rule.HasChildren) throw Invalid("a misplaced color scale, data bar or icon set");
                    BeginVisual(record);
                    return true;
                case W.BrtCFVO:
                    if (_visual == null) throw Invalid("a BrtCFVO record outside a color scale, data bar or icon set");
                    _values.Add(ReadValueObject(record));
                    return true;
                case W.BrtColor:
                    if (_visual == null) throw Invalid("a BrtColor record outside a color scale or data bar");
                    var color = new DocumentFormat.OpenXml.Spreadsheet.Color();
                    XlsbStylesheetProjector.ApplyColor(color, XlsbStylesheetReader.ReadColor(new XlsbBinaryCursor(record.Data)));
                    _colors.Add(color);
                    return true;
                case W.BrtEndColorScale:
                case W.BrtEndDatabar:
                case W.BrtEndIconSet:
                    if (_visual == null || record.Type != _visualEnd) throw Invalid("an unmatched color scale, data bar or icon set end record");
                    EndVisual(record);
                    return true;
                case W.BrtEndConditionalFormatting:
                    if (_rule != null) throw Invalid("a conditional formatting collection that ends inside a rule");
                    EndCollection(record);
                    return true;
                case BrtFrtBegin:
                    // BrtCFRuleExt and other future records link this rule to Office 2010 content.
                    _futureDepth = 1;
                    MarkUnprojected(record);
                    return true;
                default:
                    MarkUnprojected(record);
                    return true;
            }
        }

        internal void Complete() {
            if (_collection != null) throw Invalid("an unterminated conditional formatting collection");
        }

        private void BeginCollection(XlsbRecord record) {
            if (record.Data.Length < 12) throw Invalid("a truncated BrtBeginConditionalFormatting record");
            _declaredRules = BitConverter.ToInt32(record.Data, 0);
            bool pivot = BitConverter.ToUInt32(record.Data, 4) != 0;
            int offset = 8;
            _collection = new ConditionalFormatting {
                SequenceOfReferences = new ListValue<StringValue> {
                    InnerText = XlsbWorksheetFeatureReader.ReadSqRfX(record.Data, ref offset, _owner.PartName)
                }
            };
            if (pivot) _collection.Pivot = true;
            _unprojected = false;
            _unprojectedRecord = null;
        }

        private void EndCollection(XlsbRecord record) {
            ConditionalFormatting collection = _collection!;
            _collection = null;
            if (!_unprojected && collection.Elements<ConditionalFormattingRule>().Count() != _declaredRules) MarkUnprojected(record);
            if (_unprojected) {
                _owner.Sheet.HasUnprojectedConditionalFormatting = true;
                XlsbRecord source = _unprojectedRecord ?? record;
                _owner.Workbook.AddDiagnostic(new XlsbImportDiagnostic(
                    "xlsb-conditional-formatting-preserved",
                    XlsbImportDiagnosticSeverity.Warning,
                    "Preserved an XLSB conditional formatting collection that is not projected into the editable worksheet model.",
                    _owner.PartName,
                    source.Type,
                    source.Offset));
                return;
            }
            _owner.Sheet.AddFeatureElement(collection);
        }

        private ConditionalFormattingRule ReadRule(XlsbRecord record) {
            var cursor = new XlsbBinaryCursor(record.Data);
            try {
                uint type = cursor.ReadUInt32();
                uint template = cursor.ReadUInt32();
                uint differentialFormat = cursor.ReadUInt32();
                int priority = cursor.ReadInt32();
                uint parameter = cursor.ReadUInt32();
                cursor.Skip(8);
                ushort flags = cursor.ReadUInt16();
                uint[] formulaSizes = { cursor.ReadUInt32(), cursor.ReadUInt32(), cursor.ReadUInt32() };
                uint textLength = cursor.ReadUInt32();
                string? text = null;
                if (textLength != 0xFFFFFFFFU) {
                    if (textLength > 255U) throw Invalid("a BrtBeginCFRule text longer than 255 characters");
                    text = System.Text.Encoding.Unicode.GetString(cursor.ReadBytes((int)textLength * 2));
                }
                int offset = cursor.Position;
                var formulas = new List<string?>();
                (int Row, int Column) anchor = Anchor();
                foreach (uint size in formulaSizes) {
                    if (size != 0) formulas.Add(_owner.ReadFormula(record.Data, ref offset, anchor));
                }
                if (offset != record.Data.Length) throw Invalid("a BrtBeginCFRule record with trailing data");

                var rule = new ConditionalFormattingRule();
                bool projected = Project(rule, type, template, parameter, flags, text, formulaSizes);
                if (differentialFormat != 0xFFFFFFFFU && !IsVisual(rule)) rule.FormatId = differentialFormat;
                rule.Priority = priority;
                if ((flags & 0x0002) != 0) rule.StopIfTrue = true;
                // rgce3 of a visual rule has no SpreadsheetML equivalent.
                if (!projected || formulas.Any(formula => formula == null) || formulaSizes[2] != 0) {
                    MarkUnprojected(record);
                } else {
                    foreach (string? formula in formulas) rule.Append(new Formula(formula!));
                }
                return rule;
            } catch (EndOfStreamException exception) {
                throw new InvalidDataException($"The BrtBeginCFRule record at offset {record.Offset} in '{_owner.PartName}' is truncated.", exception);
            }
        }

        // Applies the type-specific attributes; false when the type/template pair or its formulas have no projection.
        private static bool Project(ConditionalFormattingRule rule, uint type, uint template, uint parameter, ushort flags, string? text, uint[] sizes) {
            int formulas = sizes.Take(2).Count(size => size != 0);
            switch (type) {
                case W.TypeCellIs when template == W.TemplateExpression && parameter >= 1 && parameter <= 8:
                    rule.Type = ConditionalFormatValues.CellIs;
                    rule.Operator = W.CellOperators[parameter - 1];
                    return formulas == (parameter <= 2 ? 2 : 1);
                case W.TypeExpression:
                    return ProjectExpression(rule, template, parameter, flags, text, formulas);
                case W.TypeFilter when template == W.TemplateFilter:
                    rule.Type = ConditionalFormatValues.Top10;
                    if ((flags & 0x0010) != 0) rule.Percent = true;
                    if ((flags & 0x0008) != 0) rule.Bottom = true;
                    rule.Rank = parameter;
                    return formulas == 0;
                case W.TypeGradient when template == W.TemplateGradient:
                    rule.Type = ConditionalFormatValues.ColorScale;
                    return formulas == 0;
                case W.TypeDataBar when template == W.TemplateDataBar:
                    rule.Type = ConditionalFormatValues.DataBar;
                    return formulas == 0;
                case W.TypeMultiState when template == W.TemplateMultiState:
                    rule.Type = ConditionalFormatValues.IconSet;
                    return formulas == 0;
                default:
                    return false;
            }
        }

        private static bool ProjectExpression(ConditionalFormattingRule rule, uint template, uint parameter, ushort flags, string? text, int formulas) {
            switch (template) {
                case W.TemplateFormula:
                    rule.Type = ConditionalFormatValues.Expression;
                    return formulas == 1;
                case W.TemplateContainsText when parameter < 4 && !string.IsNullOrEmpty(text):
                    rule.Type = W.TextOperators[parameter].Type;
                    rule.Operator = W.TextOperators[parameter].Operator;
                    rule.Text = text;
                    return formulas == 1;
                case W.TemplateContainsBlanks:
                    rule.Type = ConditionalFormatValues.ContainsBlanks;
                    return formulas == 1;
                case W.TemplateContainsNoBlanks:
                    rule.Type = ConditionalFormatValues.NotContainsBlanks;
                    return formulas == 1;
                case W.TemplateContainsErrors:
                    rule.Type = ConditionalFormatValues.ContainsErrors;
                    return formulas == 1;
                case W.TemplateContainsNoErrors:
                    rule.Type = ConditionalFormatValues.NotContainsErrors;
                    return formulas == 1;
                case W.TemplateUniqueValues:
                    rule.Type = ConditionalFormatValues.UniqueValues;
                    return formulas == 0;
                case W.TemplateDuplicateValues:
                    rule.Type = ConditionalFormatValues.DuplicateValues;
                    return formulas == 0;
                case W.TemplateAboveAverage:
                case W.TemplateBelowAverage:
                case W.TemplateEqualAboveAverage:
                case W.TemplateEqualBelowAverage:
                    rule.Type = ConditionalFormatValues.AboveAverage;
                    if (template is W.TemplateBelowAverage or W.TemplateEqualBelowAverage) rule.AboveAverage = false;
                    if (template is W.TemplateEqualAboveAverage or W.TemplateEqualBelowAverage) rule.EqualAverage = true;
                    if (parameter > 3) return false;
                    if (parameter != 0) rule.StdDev = (int)parameter;
                    return formulas == 0;
                default:
                    int period = Array.FindIndex(W.TimePeriods, item => item.Template == template);
                    if (period < 0 || W.TimePeriods[period].Parameter != parameter) return false;
                    rule.Type = ConditionalFormatValues.TimePeriod;
                    rule.TimePeriod = W.TimePeriods[period].Period;
                    return formulas == 1;
            }
        }

        // The top-left cell of the first range; CFParsedFormula PtgRefN offsets are relative to it.
        private (int Row, int Column) Anchor() {
            string first = _collection!.SequenceOfReferences!.InnerText!.Split(' ')[0];
            Write.XlsbPayloadWriter.TryParseRange(first, out int row, out int column, out _, out _);
            return (row - 1, column - 1);
        }

        private static bool IsVisual(ConditionalFormattingRule rule) {
            ConditionalFormatValues? type = rule.Type?.Value;
            return type == ConditionalFormatValues.ColorScale || type == ConditionalFormatValues.DataBar || type == ConditionalFormatValues.IconSet;
        }

        private void BeginVisual(XlsbRecord record) {
            _values.Clear();
            _colors.Clear();
            switch (record.Type) {
                case W.BrtBeginColorScale:
                    _visual = new ColorScale();
                    _visualEnd = W.BrtEndColorScale;
                    break;
                case W.BrtBeginDatabar: {
                    if (record.Data.Length < 3) throw Invalid("a truncated BrtBeginDatabar record");
                    var dataBar = new DataBar();
                    if (record.Data[0] != 10) dataBar.MinLength = record.Data[0];
                    if (record.Data[1] != 90) dataBar.MaxLength = record.Data[1];
                    if (record.Data[2] == 0) dataBar.ShowValue = false;
                    _visual = dataBar;
                    _visualEnd = W.BrtEndDatabar;
                    break;
                }
                default: {
                    if (record.Data.Length < 6) throw Invalid("a truncated BrtBeginIconSet record");
                    uint set = BitConverter.ToUInt32(record.Data, 0);
                    ushort flags = BitConverter.ToUInt16(record.Data, 4);
                    var iconSet = new IconSet();
                    if (set < W.IconSets.Length) iconSet.IconSetValue = W.IconSets[set];
                    else MarkUnprojected(record);
                    if ((flags & 0x0002) != 0) iconSet.ShowValue = false;
                    if ((flags & 0x0004) != 0) iconSet.Reverse = true;
                    _visual = iconSet;
                    _visualEnd = W.BrtEndIconSet;
                    break;
                }
            }
        }

        private void EndVisual(XlsbRecord record) {
            OpenXmlCompositeElement visual = _visual!;
            _visual = null;
            bool valid = visual switch {
                ColorScale => _values.Count is 2 or 3 && _colors.Count == _values.Count,
                DataBar => _values.Count == 2 && _colors.Count == 1,
                _ => _values.Count is >= 3 and <= 5 && _colors.Count == 0,
            };
            ConditionalFormatValues? expected = visual switch {
                ColorScale => ConditionalFormatValues.ColorScale,
                DataBar => ConditionalFormatValues.DataBar,
                _ => ConditionalFormatValues.IconSet,
            };
            if (!valid || _rule!.Type?.Value != expected) {
                MarkUnprojected(record);
                return;
            }
            foreach (ConditionalFormatValueObject value in _values) visual.Append(value);
            foreach (DocumentFormat.OpenXml.Spreadsheet.Color color in _colors) visual.Append(color);
            _rule.Append(visual);
        }

        private ConditionalFormatValueObject ReadValueObject(XlsbRecord record) {
            var cursor = new XlsbBinaryCursor(record.Data);
            try {
                uint type = cursor.ReadUInt32();
                double number = cursor.ReadDouble();
                cursor.ReadUInt32(); // fSaveGTE
                bool greaterOrEqual = cursor.ReadUInt32() != 0;
                uint formulaSize = cursor.ReadUInt32();
                string? formula = null;
                // MS-XLSB 2.4.334: cbFmla and the formula are ignored for CFVOMIN and CFVOMAX thresholds.
                if (formulaSize != 0 && type is not (2 or 3)) {
                    int offset = cursor.Position;
                    formula = _owner.ReadFormula(record.Data, ref offset);
                    if (formula == null) MarkUnprojected(record);
                }
                int index = Array.FindIndex(W.ValueObjectTypes, item => item.Value == type);
                var value = new ConditionalFormatValueObject();
                if (index < 0) {
                    MarkUnprojected(record);
                    return value;
                }
                value.Type = W.ValueObjectTypes[index].Type;
                if (formula != null) value.Val = formula;
                else if (type is 1 or 4 or 5) value.Val = InvariantNumberText.Get(number);
                if (_visual is IconSet && !greaterOrEqual) value.GreaterThanOrEqual = false;
                return value;
            } catch (EndOfStreamException exception) {
                throw new InvalidDataException($"The BrtCFVO record at offset {record.Offset} in '{_owner.PartName}' is truncated.", exception);
            }
        }

        private void MarkUnprojected(XlsbRecord record) {
            if (_unprojected) return;
            _unprojected = true;
            _unprojectedRecord = record;
        }

        private InvalidDataException Invalid(string description) =>
            new InvalidDataException($"The XLSB worksheet part '{_owner.PartName}' contains {description}.");
    }
}
