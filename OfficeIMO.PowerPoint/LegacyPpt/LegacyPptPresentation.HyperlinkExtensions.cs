using OfficeIMO.PowerPoint.LegacyPpt.Diagnostics;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt {
    public sealed partial class LegacyPptPresentation {
        private const ushort RecordDocumentInfoList = 0x07D0;
        private const ushort RecordExternalHyperlink9 = 0x0FE4;
        private const ushort RecordExternalHyperlinkFlagsAtom = 0x1018;
        private const string Ppt9TagName = "___PPT9";

        private void ParseHyperlinkExtensions(LegacyPptRecord document,
            LegacyPptImportOptions options) {
            LegacyPptRecord[] infoLists = document.Children.Where(record =>
                record.Type == RecordDocumentInfoList).ToArray();
            if (infoLists.Length > 1) {
                AddDiagnostic("PPT-HYPERLINK9-DOCINFO",
                    LegacyPptDiagnosticSeverity.Warning,
                    "The document has multiple information lists; hyperlink extensions remain preserve-only.",
                    infoLists[1].Offset);
                return;
            }
            if (infoLists.Length == 0) return;

            var seenIds = new HashSet<uint>();
            foreach (LegacyPptRecord progTags in infoLists[0].Children.Where(record =>
                         record.Type == RecordProgTags)) {
                foreach (LegacyPptRecord binaryTag in progTags.Children.Where(record =>
                             record.Type == RecordProgBinaryTag)) {
                    if (!IsPpt9BinaryTag(binaryTag)) continue;
                    LegacyPptRecord[] dataBlobs = binaryTag.Children.Where(record =>
                        record.Type == RecordBinaryTagDataBlob).ToArray();
                    if (dataBlobs.Length != 1 || dataBlobs[0].Version != 0
                        || dataBlobs[0].Instance != 0) {
                        AddDiagnostic("PPT-HYPERLINK9-DATA",
                            LegacyPptDiagnosticSeverity.Warning,
                            "A PP9 programmable tag has no unique valid data blob; hyperlink extensions remain preserve-only.",
                            binaryTag.Offset);
                        continue;
                    }
                    IReadOnlyList<LegacyPptRecord> records;
                    try {
                        byte[] bytes = dataBlobs[0].CopyRecordBytes();
                        records = LegacyPptRecordReader.ReadSequence(bytes, 8,
                            dataBlobs[0].PayloadLength, options,
                            _recordBudget);
                    } catch (InvalidDataException) {
                        AddDiagnostic("PPT-HYPERLINK9-TRUNCATED",
                            LegacyPptDiagnosticSeverity.Warning,
                            "A PP9 programmable tag is malformed; hyperlink extensions remain preserve-only.",
                            dataBlobs[0].Offset);
                        continue;
                    }
                    foreach (LegacyPptRecord extension in records.Where(record =>
                                 record.Type == RecordExternalHyperlink9)) {
                        ApplyHyperlinkExtension(extension, seenIds, options);
                    }
                }
            }
        }

        private bool IsPpt9BinaryTag(LegacyPptRecord record) {
            if (record.Version != 0x0F || record.Instance != 0
                || record.Type != RecordProgBinaryTag) return false;
            LegacyPptRecord[] names = record.Children.Where(child =>
                child.Type == RecordCString && child.Instance == 0).ToArray();
            return names.Length == 1 && TryReadUnicodeString(names[0], out string name)
                && string.Equals(name, Ppt9TagName, StringComparison.Ordinal);
        }

        private void ApplyHyperlinkExtension(LegacyPptRecord container,
            ISet<uint> seenIds, LegacyPptImportOptions options) {
            if (container.Version != 0x0F || container.Instance != 0) {
                AddDiagnostic("PPT-HYPERLINK9-CONTAINER",
                    LegacyPptDiagnosticSeverity.Warning,
                    "An ExHyperlink9Container has an invalid header and remains preserve-only.",
                    container.Offset);
                return;
            }
            LegacyPptRecord[] references = container.Children.Where(record =>
                record.Type == RecordExternalHyperlinkAtom).ToArray();
            LegacyPptRecord[] flagsAtoms = container.Children.Where(record =>
                record.Type == RecordExternalHyperlinkFlagsAtom).ToArray();
            LegacyPptRecord[] screenTips = container.Children.Where(record =>
                record.Type == RecordCString).ToArray();
            if (references.Length != 1 || references[0].Version != 0
                || references[0].Instance != 0 || references[0].PayloadLength != 4
                || flagsAtoms.Length != 1 || flagsAtoms[0].Version != 0
                || flagsAtoms[0].Instance != 0 || flagsAtoms[0].PayloadLength != 4
                || screenTips.Length > 1 || screenTips.Any(record =>
                    record.Version != 0 || record.Instance != 0
                    || (record.PayloadLength & 1) != 0)) {
                AddDiagnostic("PPT-HYPERLINK9-SHAPE",
                    LegacyPptDiagnosticSeverity.Warning,
                    "An ExHyperlink9Container has malformed reference, screen-tip, or flags data and remains preserve-only.",
                    container.Offset);
                return;
            }
            uint id = references[0].ReadUInt32(0);
            if (!seenIds.Add(id)) {
                AddDiagnostic("PPT-HYPERLINK9-DUPLICATE",
                    LegacyPptDiagnosticSeverity.Warning,
                    $"Hyperlink identifier {id} has multiple PP9 extensions and remains preserve-only.",
                    container.Offset);
                return;
            }
            if (!_hyperlinksById.TryGetValue(id, out LegacyPptHyperlink? hyperlink)) {
                AddDiagnostic("PPT-HYPERLINK9-MISSING",
                    LegacyPptDiagnosticSeverity.Warning,
                    $"A PP9 extension references missing hyperlink identifier {id}.",
                    container.Offset);
                return;
            }
            string? screenTip = screenTips.Length == 0
                ? null
                : screenTips[0].ReadUtf16Text().TrimEnd('\0');
            uint flags = flagsAtoms[0].ReadUInt32(0);
            hyperlink.ApplyExtension(screenTip, flags);
            if (options.ReportUnsupportedContent && (flags & ~0x07U) != 0) {
                AddDiagnostic("PPT-HYPERLINK9-RESERVED",
                    LegacyPptDiagnosticSeverity.Warning,
                    $"Hyperlink identifier {id} has nonzero reserved PP9 flags that remain preserve-only.",
                    flagsAtoms[0].Offset);
            }
            if (options.ReportUnsupportedContent && flags != 0) {
                AddDiagnostic("PPT-HYPERLINK9-FLAGS",
                    LegacyPptDiagnosticSeverity.Warning,
                    $"Hyperlink identifier {id} has PP9 dialog or custom-show flags that remain preserve-only.",
                    flagsAtoms[0].Offset);
            }
        }
    }
}
