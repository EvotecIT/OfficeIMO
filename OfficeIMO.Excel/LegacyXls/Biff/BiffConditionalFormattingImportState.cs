using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal sealed class BiffConditionalFormattingImportState {
        private readonly LegacyXlsWorkbook _workbook;
        private readonly LegacyXlsWorksheet _sheet;
        private readonly IReadOnlyList<BiffExternSheetReference> _externSheets;
        private readonly IReadOnlyList<LegacyXlsExternalReference> _externalReferences;
        private readonly IReadOnlyList<string> _sheetNames;
        private readonly IReadOnlyList<string?> _definedNames;
        private readonly IReadOnlyList<LegacyXlsDifferentialFormat> _differentialFormats;
        private readonly Dictionary<ushort, List<LegacyXlsConditionalFormatting>> _rulesByHeaderId = new();
        private IReadOnlyList<string> _ranges = Array.Empty<string>();
        private ushort _headerId;
        private ushort _expectedRuleCount;
        private ushort _readRuleCount;
        private int _headerOffset;
        private int _headerPayloadLength;

        internal BiffConditionalFormattingImportState(
            LegacyXlsWorkbook workbook,
            LegacyXlsWorksheet sheet,
            IReadOnlyList<BiffExternSheetReference> externSheets,
            IReadOnlyList<LegacyXlsExternalReference> externalReferences,
            IReadOnlyList<string> sheetNames,
            IReadOnlyList<string?> definedNames,
            IReadOnlyList<LegacyXlsDifferentialFormat> differentialFormats) {
            _workbook = workbook;
            _sheet = sheet;
            _externSheets = externSheets;
            _externalReferences = externalReferences;
            _sheetNames = sheetNames;
            _definedNames = definedNames;
            _differentialFormats = differentialFormats;
        }

        internal bool HasPendingHeader => _expectedRuleCount > _readRuleCount;

        internal bool TryReadHeader(byte[] payload, int recordOffset) {
            if (!BiffConditionalFormattingReader.TryReadHeader(payload, out ushort ruleCount, out ushort headerId, out IReadOnlyList<string> ranges)) {
                Clear();
                return false;
            }

            _headerId = headerId;
            _expectedRuleCount = ruleCount;
            _readRuleCount = 0;
            _ranges = ranges;
            _headerOffset = recordOffset;
            _headerPayloadLength = payload.Length;
            return true;
        }

        internal bool TryReadRule(byte[] payload) {
            return TryReadRule(payload, recordOffset: 0, out _);
        }

        internal bool TryReadRule(byte[] payload, int recordOffset, out BiffFormulaReadFailure? formulaFailure) {
            formulaFailure = null;
            if (!HasPendingHeader) {
                return false;
            }

            bool parsed = BiffConditionalFormattingReader.TryReadRule(
                payload,
                _workbook,
                recordOffset,
                _externSheets,
                _externalReferences,
                _sheetNames,
                _definedNames,
                _ranges,
                out LegacyXlsConditionalFormatting? conditionalFormatting,
                out formulaFailure,
                out bool hasInlineFormatting);

            if (parsed) {
                LegacyXlsConditionalFormatting rule = conditionalFormatting!;
                _sheet.AddConditionalFormatting(rule);
                if (!_rulesByHeaderId.TryGetValue(_headerId, out List<LegacyXlsConditionalFormatting>? rules)) {
                    rules = new List<LegacyXlsConditionalFormatting>();
                    _rulesByHeaderId[_headerId] = rules;
                }

                rules.Add(rule);
            }

            _readRuleCount++;
            if (!HasPendingHeader) {
                Clear();
            }

            return parsed && !hasInlineFormatting;
        }

        internal bool TryReadExtension(byte[] payload, int recordOffset, out bool hasUnprojectedFormatting) {
            hasUnprojectedFormatting = false;
            if (!BiffConditionalFormattingReader.TryReadExtension(
                payload,
                recordOffset,
                _sheet.Name,
                matchedRule: false,
                out LegacyXlsConditionalFormattingExtensionRecord? extensionRecord,
                out ushort headerId,
                out ushort ruleIndex,
                out int? priority,
                out bool stopIfTrue,
                out bool hasInlineFormatting)) {
                if (extensionRecord != null) {
                    _sheet.AddConditionalFormattingExtension(extensionRecord);
                }

                hasUnprojectedFormatting = hasInlineFormatting;
                return false;
            }

            if (!_rulesByHeaderId.TryGetValue(headerId, out List<LegacyXlsConditionalFormatting>? rules)) {
                if (_rulesByHeaderId.Count != 1) {
                    if (extensionRecord != null) {
                        _sheet.AddConditionalFormattingExtension(extensionRecord);
                    }

                    hasUnprojectedFormatting = hasInlineFormatting;
                    return false;
                }

                foreach (List<LegacyXlsConditionalFormatting> onlyRules in _rulesByHeaderId.Values) {
                    rules = onlyRules;
                    break;
                }
            }

            if (rules == null || ruleIndex >= rules.Count) {
                if (extensionRecord != null) {
                    _sheet.AddConditionalFormattingExtension(extensionRecord);
                }

                hasUnprojectedFormatting = hasInlineFormatting;
                return false;
            }

            extensionRecord?.MarkMatchedRule();
            if (extensionRecord != null) {
                _sheet.AddConditionalFormattingExtension(extensionRecord);
            }

            LegacyXlsDifferentialFormat? differentialFormat = hasInlineFormatting
                ? ResolveDifferentialFormat(headerId)
                : null;
            if (differentialFormat != null) {
                extensionRecord?.MarkProjectedFormatting();
            }

            rules[ruleIndex].ApplyExtension(priority, stopIfTrue, differentialFormat);
            hasUnprojectedFormatting = hasInlineFormatting && differentialFormat == null;
            return true;
        }

        private LegacyXlsDifferentialFormat? ResolveDifferentialFormat(ushort headerId) {
            if (headerId > 0 && headerId <= _differentialFormats.Count) {
                return _differentialFormats[headerId - 1];
            }

            return _differentialFormats.Count == 1 ? _differentialFormats[0] : null;
        }

        internal void AddUnresolvedFeatures(
            List<LegacyXlsUnsupportedFeature> unsupportedFeatures,
            List<LegacyXlsPreservedFeatureRecord> preservedFeatureRecords,
            List<LegacyXlsImportDiagnostic> diagnostics,
            bool reportUnsupportedRecords) {
            if (!HasPendingHeader) {
                return;
            }

            LegacyXlsUnsupportedFeature feature = BiffUnsupportedRecordDiagnostics.CreateUnsupportedRecordFeature(
                (ushort)BiffRecordType.CondFmt,
                _headerOffset,
                _sheet.Name);
            unsupportedFeatures.Add(feature);
            if (BiffUnsupportedRecordDiagnostics.TryCreatePreservedFeatureRecord(feature, _headerPayloadLength, out LegacyXlsPreservedFeatureRecord? preservedRecord)) {
                preservedFeatureRecords.Add(preservedRecord!);
            }

            if (reportUnsupportedRecords) {
                BiffUnsupportedRecordDiagnostics.AddUnsupportedRecordDiagnostic(
                    diagnostics,
                    (ushort)BiffRecordType.CondFmt,
                    _headerOffset,
                    _sheet.Name);
            }

            Clear();
        }

        internal void Clear() {
            _headerId = 0;
            _expectedRuleCount = 0;
            _readRuleCount = 0;
            _headerOffset = 0;
            _headerPayloadLength = 0;
            _ranges = Array.Empty<string>();
        }
    }
}
