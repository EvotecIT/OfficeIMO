using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal sealed class BiffConditionalFormattingImportState {
        private readonly LegacyXlsWorksheet _sheet;
        private readonly IReadOnlyList<BiffExternSheetReference> _externSheets;
        private readonly IReadOnlyList<LegacyXlsExternalReference> _externalReferences;
        private readonly IReadOnlyList<string> _sheetNames;
        private readonly IReadOnlyList<string?> _definedNames;
        private IReadOnlyList<string> _ranges = Array.Empty<string>();
        private ushort _expectedRuleCount;
        private ushort _readRuleCount;
        private int _headerOffset;
        private int _headerPayloadLength;

        internal BiffConditionalFormattingImportState(
            LegacyXlsWorksheet sheet,
            IReadOnlyList<BiffExternSheetReference> externSheets,
            IReadOnlyList<LegacyXlsExternalReference> externalReferences,
            IReadOnlyList<string> sheetNames,
            IReadOnlyList<string?> definedNames) {
            _sheet = sheet;
            _externSheets = externSheets;
            _externalReferences = externalReferences;
            _sheetNames = sheetNames;
            _definedNames = definedNames;
        }

        internal bool HasPendingHeader => _expectedRuleCount > _readRuleCount;

        internal bool TryReadHeader(byte[] payload, int recordOffset) {
            if (!BiffConditionalFormattingReader.TryReadHeader(payload, out ushort ruleCount, out IReadOnlyList<string> ranges)) {
                Clear();
                return false;
            }

            _expectedRuleCount = ruleCount;
            _readRuleCount = 0;
            _ranges = ranges;
            _headerOffset = recordOffset;
            _headerPayloadLength = payload.Length;
            return true;
        }

        internal bool TryReadRule(byte[] payload) {
            if (!HasPendingHeader) {
                return false;
            }

            bool parsed = BiffConditionalFormattingReader.TryReadRule(
                payload,
                _externSheets,
                _externalReferences,
                _sheetNames,
                _definedNames,
                _ranges,
                out LegacyXlsConditionalFormatting? conditionalFormatting);

            if (parsed) {
                _sheet.AddConditionalFormatting(conditionalFormatting!);
            }

            _readRuleCount++;
            if (!HasPendingHeader) {
                Clear();
            }

            return parsed;
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
            _expectedRuleCount = 0;
            _readRuleCount = 0;
            _headerOffset = 0;
            _headerPayloadLength = 0;
            _ranges = Array.Empty<string>();
        }
    }
}
