using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Model;
using W = OfficeIMO.Excel.Xlsb.Write.XlsbWorksheetDataValidationWriter;

namespace OfficeIMO.Excel.Xlsb.Read {
    /// <summary>
    /// Decodes a DVALS collection (MS-XLSB 2.4.55, 2.4.356) into dataValidations. The optional alternate-content blocks
    /// inside it are consumed too: ACDVALLIST (BrtDValList) replaces formula1 of the next BrtDVal, and ACUID revision
    /// identifiers are not part of the model.
    /// </summary>
    internal sealed class XlsbWorksheetDataValidationReader {
        private readonly XlsbWorksheetFeatureReader _owner;
        private DataValidations? _current;
        private bool _seen;
        private bool _undecoded;
        private int _alternateContentDepth;
        private string? _pendingList;

        internal XlsbWorksheetDataValidationReader(XlsbWorksheetFeatureReader owner) => _owner = owner;

        internal bool TryRead(XlsbRecord record) {
            switch (record.Type) {
                case W.BrtBeginDVals:
                    if (_current != null || _seen) {
                        throw new InvalidDataException($"The XLSB worksheet part '{_owner.PartName}' contains a repeated data-validation collection.");
                    }
                    if (record.Data.Length < 18) {
                        throw new InvalidDataException($"The XLSB worksheet part '{_owner.PartName}' contains a truncated BrtBeginDVals record.");
                    }
                    _seen = true;
                    _current = new DataValidations();
                    if ((record.Data[0] & 0x01) != 0) _current.DisablePrompts = true;
                    uint xLeft = BitConverter.ToUInt32(record.Data, 2);
                    uint yTop = BitConverter.ToUInt32(record.Data, 6);
                    if (xLeft != 0) _current.XWindow = xLeft;
                    if (yTop != 0) _current.YWindow = yTop;
                    return true;
                // A BrtDVal outside DVALS breaks the WORKSHEET grammar; it stays a preserved record as before.
                case W.BrtDVal when _current != null:
                    DataValidation? validation = ReadValidation(record.Data, _pendingList);
                    _pendingList = null;
                    if (validation == null) {
                        if (!_undecoded) {
                            _owner.Workbook.AddDiagnostic(new XlsbImportDiagnostic(
                                "xlsb-data-validation-undecoded",
                                XlsbImportDiagnosticSeverity.Warning,
                                $"Worksheet '{_owner.Sheet.Name}' has data validations that OfficeIMO cannot decode; they are kept only while they stay unchanged in a native XLSB save.",
                                _owner.PartName, record.Type, record.Offset));
                        }
                        _undecoded = true;
                    } else {
                        _current.Append(validation);
                    }
                    return true;
                case W.BrtACBegin when _current != null:
                    _alternateContentDepth++;
                    return true;
                case W.BrtACEnd when _current != null:
                    if (_alternateContentDepth == 0) {
                        throw new InvalidDataException($"The XLSB worksheet part '{_owner.PartName}' contains an unmatched BrtACEnd record.");
                    }
                    _alternateContentDepth--;
                    return true;
                case W.BrtUid when _current != null && _alternateContentDepth > 0:
                    return true;
                case W.BrtDValList when _current != null && _alternateContentDepth > 0:
                    int offset = 0;
                    _pendingList = ReadWideString(record.Data, ref offset);
                    return true;
                case W.BrtEndDVals:
                    if (_current == null || _alternateContentDepth != 0) {
                        throw new InvalidDataException($"The XLSB worksheet part '{_owner.PartName}' contains an unmatched data-validation end record.");
                    }
                    int count = _current.Elements<DataValidation>().Count();
                    if (_undecoded) {
                        _owner.Sheet.HasUndecodedDataValidations = true;
                    } else if (count != 0) {
                        // OfficeIMO states the derived count on save; project it the same way so unchanged content compares equal.
                        _current.Count = (uint)count;
                        _owner.Sheet.AddFeatureElement(_current);
                    }
                    _current = null;
                    return true;
                default:
                    return false;
            }
        }

        internal void Complete() {
            if (_current != null) {
                throw new InvalidDataException($"The XLSB worksheet part '{_owner.PartName}' contains an unterminated data-validation collection.");
            }
        }

        /// <summary>Decodes one BrtDVal; null when a value or formula has no SpreadsheetML form OfficeIMO can produce.</summary>
        private DataValidation? ReadValidation(byte[] data, string? listFormula) {
            if (data.Length < 4) throw new InvalidDataException($"The XLSB worksheet part '{_owner.PartName}' contains a truncated BrtDVal record.");
            uint flags = BitConverter.ToUInt32(data, 0);
            int valType = (int)(flags & 0x0F);
            int errStyle = (int)((flags >> 4) & 0x07);
            int imeMode = (int)((flags >> 10) & 0xFF);
            int typOperator = (int)((flags >> 20) & 0x0F);
            int offset = 4;
            string references = XlsbWorksheetFeatureReader.ReadSqRfX(data, ref offset, _owner.PartName);
            string? errorTitle = ReadNullableWideString(data, ref offset);
            string? error = ReadNullableWideString(data, ref offset);
            string? promptTitle = ReadNullableWideString(data, ref offset);
            string? prompt = ReadNullableWideString(data, ref offset);
            if (references.Length == 0) return null;
            int anchorOffset = 8;
            XlsbCellRange first = XlsbWorksheetFeatureReader.ReadRfX(data, ref anchorOffset, _owner.PartName);
            var anchor = (Row: first.FirstRow - 1, Column: first.FirstColumn - 1);
            string? formula1 = _owner.ReadFormula(data, ref offset, anchor);
            string? formula2 = _owner.ReadFormula(data, ref offset, anchor);
            if (listFormula != null && valType == 3) formula1 = listFormula;

            if (valType >= W.Types.Length || errStyle >= W.ErrorStyles.Length || imeMode >= W.ImeModes.Length
                || !W.IgnoresOperator(valType) && typOperator >= W.Operators.Length
                || formula1 == null || formula2 == null) {
                return null;
            }

            var validation = new DataValidation();
            if (valType != 0) validation.Type = W.Types[valType];
            if (errStyle != 0) validation.ErrorStyle = W.ErrorStyles[errStyle];
            if (imeMode != 0) validation.ImeMode = W.ImeModes[imeMode];
            if (!W.IgnoresOperator(valType)) validation.Operator = W.Operators[typOperator];
            if ((flags & W.AllowBlankFlag) != 0) validation.AllowBlank = true;
            if ((flags & W.SuppressComboFlag) != 0) validation.ShowDropDown = true;
            if ((flags & W.ShowInputMessageFlag) != 0) validation.ShowInputMessage = true;
            if ((flags & W.ShowErrorMessageFlag) != 0) validation.ShowErrorMessage = true;
            if (errorTitle != null) validation.ErrorTitle = errorTitle;
            if (error != null) validation.Error = error;
            if (promptTitle != null) validation.PromptTitle = promptTitle;
            if (prompt != null) validation.Prompt = prompt;
            validation.SequenceOfReferences = new ListValue<StringValue> { InnerText = references };
            // formula1 is ignored for "any value" and formula2 is ignored unless the operator takes two bounds.
            if (valType != 0 && formula1.Length != 0) validation.Append(new Formula1(formula1));
            if (!W.IgnoresOperator(valType) && typOperator <= 1 && formula2.Length != 0) validation.Append(new Formula2(formula2));
            return validation;
        }

        private string? ReadNullableWideString(byte[] data, ref int offset) {
            if (offset > data.Length - 4) throw new InvalidDataException($"A string in '{_owner.PartName}' is truncated.");
            if (BitConverter.ToUInt32(data, offset) == 0xFFFFFFFFU) {
                offset += 4;
                return null;
            }
            return ReadWideString(data, ref offset);
        }

        private string ReadWideString(byte[] data, ref int offset) {
            if (offset > data.Length - 4) throw new InvalidDataException($"A string in '{_owner.PartName}' is truncated.");
            uint length = BitConverter.ToUInt32(data, offset);
            offset += 4;
            if (length > (uint)(data.Length - offset) / 2) throw new InvalidDataException($"A string in '{_owner.PartName}' is truncated.");
            string value = System.Text.Encoding.Unicode.GetString(data, offset, (int)length * 2);
            offset += (int)length * 2;
            return value;
        }
    }
}
