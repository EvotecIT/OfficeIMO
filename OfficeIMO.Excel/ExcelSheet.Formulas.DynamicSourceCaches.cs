using DocumentFormat.OpenXml.Spreadsheet;
using System.Security.Cryptography;
using System.Xml;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private const long MaxOriginalDynamicSpillXmlBytes = 64L * 1024 * 1024;
        private const string SpreadsheetMainNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private const string StrictSpreadsheetMainNamespace = "http://purl.oclc.org/ooxml/spreadsheetml/main";

        private void CaptureOriginalDynamicSpillFingerprintIfSafe() {
            // Workbook-only edits leave this part untouched. Once the raw package
            // is exposed or its root loaded, prior external edits are ambiguous.
            if (!_excelDocument.IsOpenXmlDocumentExposed && !_worksheetPart.IsRootElementLoaded)
                CaptureOriginalDynamicSpillFingerprint();
        }

        // A fingerprint costs no retained copy of a loaded worksheet. It also prevents a
        // later worksheet-part Save from becoming an accidental ownership baseline.
        private void CaptureOriginalDynamicSpillFingerprint() {
            DynamicSpillOwnership ownership = SpillOwnership;
            lock (ownership) {
                if (ownership.OriginalFingerprintAttempted) return;
                ownership.OriginalFingerprintAttempted = true;
                try {
                    if (!HasDynamicArrayMetadataType()) return;

                    ownership.OriginalFingerprint = ReadBoundedWorksheetFingerprint();
                } catch (Exception ex) when (ex is IOException || ex is InvalidOperationException
                    || ex is InvalidDataException || ex is XmlException) {
                    // The normal formula path still validates metadata; this optional baseline fails closed.
                }
            }
        }

        private bool HasDynamicArrayMetadataType() {
            var part = _excelDocument.WorkbookPartRoot.CellMetadataPart;
            if (part == null) return false;
            if (!part.IsRootElementLoaded) ValidateInCellImageMetadataPart(part, "Cell metadata");
            using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
            using XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                IgnoreComments = true,
                IgnoreProcessingInstructions = true,
                MaxCharactersInDocument = MaximumRichValueMetadataBytes
            });
            while (reader.Read())
                if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "metadataType"
                    && IsSpreadsheetMainNamespace(reader.NamespaceURI)
                    && string.Equals(reader.GetAttribute("name"), "XLDAPR", StringComparison.OrdinalIgnoreCase))
                    return true;
            return false;
        }

        private byte[]? ReadBoundedWorksheetFingerprint() {
            try {
                using Stream stream = _worksheetPart.GetStream(FileMode.Open, FileAccess.Read);
                if (!stream.CanSeek || stream.Length > MaxOriginalDynamicSpillXmlBytes) return null;
                using SHA256 sha = SHA256.Create();
                return sha.ComputeHash(stream);
            } catch (Exception ex) when (ex is IOException || ex is InvalidOperationException
                || ex is NotSupportedException) {
                return null;
            }
        }

        private sealed class OriginalDynamicOwnerCaches {
            internal readonly FixedArrayOwner Owner;
            internal readonly Dictionary<long, DynamicSpillCacheSnapshot> Cells =
                new Dictionary<long, DynamicSpillCacheSnapshot>();
            internal readonly HashSet<long> SeenChildren = new HashSet<long>();
            internal bool AnchorSeen;
            internal bool Valid = true;

            internal OriginalDynamicOwnerCaches(FixedArrayOwner owner) => Owner = owner;
        }

        private void TryCaptureOriginalDynamicSpillCaches() {
            DynamicSpillOwnership ownership = SpillOwnership;
            if (ownership.OriginalScanAttempted) return;
            ownership.OriginalScanAttempted = true;
            if (ownership.OriginalFingerprint == null) return;

            byte[]? currentFingerprint = ReadBoundedWorksheetFingerprint();
            if (currentFingerprint == null || !currentFingerprint.SequenceEqual(ownership.OriginalFingerprint))
                return;

            var anchors = new Dictionary<long, OriginalDynamicOwnerCaches>();
            var children = new Dictionary<long, OriginalDynamicOwnerCaches>();
            var states = new List<OriginalDynamicOwnerCaches>();
            long remainingCells = MaxResolvedFormulaRangeCells;
            foreach (FixedArrayOwner owner in GetFixedArraySheetIndex().DynamicOwners) {
                long count = (long)(owner.Bottom - owner.Top + 1) * (owner.Right - owner.Left + 1);
                if (ownership.WrittenOwners.Contains(DynamicCellKey(owner.Top, owner.Left))
                    || count < 1 || count > remainingCells) continue;
                remainingCells -= count;
                var state = new OriginalDynamicOwnerCaches(owner);
                states.Add(state);
                long anchorKey = DynamicCellKey(owner.Top, owner.Left);
                if (anchors.ContainsKey(anchorKey) || children.ContainsKey(anchorKey)) return;
                anchors.Add(anchorKey, state);
                for (int row = owner.Top; row <= owner.Bottom; row++)
                    for (int column = owner.Left; column <= owner.Right; column++) {
                        long key = DynamicCellKey(row, column);
                        if (key == anchorKey) continue;
                        if (children.ContainsKey(key) || anchors.ContainsKey(key)) return;
                        children.Add(key, state);
                    }
            }
            if (states.Count == 0 || !ReadOriginalDynamicSpillCells(anchors, children)) return;
            foreach (OriginalDynamicOwnerCaches state in states)
                if (state.Valid && state.AnchorSeen)
                    foreach (var entry in state.Cells)
                        if (!ownership.Cells.ContainsKey(entry.Key)) ownership.Cells.Add(entry.Key, entry.Value);
        }

        private bool ReadOriginalDynamicSpillCells(
            Dictionary<long, OriginalDynamicOwnerCaches> anchors,
            Dictionary<long, OriginalDynamicOwnerCaches> children) {
            try {
                using Stream stream = _worksheetPart.GetStream(FileMode.Open, FileAccess.Read);
                if (!stream.CanSeek || stream.Length > MaxOriginalDynamicSpillXmlBytes) return false;
                using XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings {
                    DtdProcessing = DtdProcessing.Prohibit,
                    XmlResolver = null,
                    IgnoreComments = true,
                    IgnoreProcessingInstructions = true,
                    MaxCharactersInDocument = MaxOriginalDynamicSpillXmlBytes
                });
                int sheetDataDepth = -1;
                while (reader.Read()) {
                    if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == sheetDataDepth
                        && reader.LocalName == "sheetData") sheetDataDepth = -1;
                    if (reader.NodeType != XmlNodeType.Element || !IsSpreadsheetMainNamespace(reader.NamespaceURI))
                        continue;
                    if (reader.LocalName == "sheetData") {
                        sheetDataDepth = reader.Depth;
                        continue;
                    }
                    if (sheetDataDepth < 0 || reader.LocalName != "c") continue;
                    if (!TryParseCellReference(reader.GetAttribute("r") ?? "", out int row, out int column))
                        continue;
                    long key = DynamicCellKey(row, column);
                    if (anchors.TryGetValue(key, out OriginalDynamicOwnerCaches? anchor)) {
                        if (anchor.AnchorSeen) anchor.Valid = false;
                        anchor.AnchorSeen = true;
                        anchor.Valid &= OriginalDynamicAnchorMatches(reader, anchor.Owner);
                        continue;
                    }
                    if (!children.TryGetValue(key, out OriginalDynamicOwnerCaches? child)) continue;
                    if (!child.SeenChildren.Add(key)) {
                        child.Valid = false;
                        continue;
                    }
                    if (TryReadOriginalPlainCache(reader, out DynamicSpillCacheSnapshot? snapshot))
                        child.Cells.Add(key, snapshot!);
                }
                return true;
            } catch (Exception ex) when (ex is IOException || ex is InvalidOperationException
                || ex is XmlException || ex is NotSupportedException) {
                return false;
            }
        }

        private static bool OriginalDynamicAnchorMatches(XmlReader reader, FixedArrayOwner owner) {
            string? cellMetadata = reader.GetAttribute("cm");
            if (!uint.TryParse(cellMetadata, out uint metadata)
                || metadata != owner.Cell.CellMetaIndex?.Value) return false;
            string expectedRange = A1.CellReference(owner.Top, owner.Left) + ":"
                + A1.CellReference(owner.Bottom, owner.Right);
            using XmlReader child = reader.ReadSubtree();
            while (child.Read()) {
                if (child.NodeType != XmlNodeType.Element || child.LocalName != "f"
                    || !IsSpreadsheetMainNamespace(child.NamespaceURI)) continue;
                string? type = child.GetAttribute("t");
                string? range = child.GetAttribute("ref");
                string formula = child.ReadElementContentAsString();
                return type == "array" && string.Equals(range, expectedRange, StringComparison.Ordinal)
                    && string.Equals(formula, owner.Cell.CellFormula?.Text, StringComparison.Ordinal);
            }
            return false;
        }

        private static bool TryReadOriginalPlainCache(XmlReader reader, out DynamicSpillCacheSnapshot? snapshot) {
            snapshot = null;
            if (reader.GetAttribute("cm") != null) return false;
            string? type = reader.GetAttribute("t");
            DocumentFormat.OpenXml.Spreadsheet.CellValues? parsedType = type switch {
                null => null,
                "b" => DocumentFormat.OpenXml.Spreadsheet.CellValues.Boolean,
                "d" => DocumentFormat.OpenXml.Spreadsheet.CellValues.Date,
                "e" => DocumentFormat.OpenXml.Spreadsheet.CellValues.Error,
                "n" => DocumentFormat.OpenXml.Spreadsheet.CellValues.Number,
                "s" => DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString,
                "str" => DocumentFormat.OpenXml.Spreadsheet.CellValues.String,
                _ => null
            };
            if (type != null && parsedType == null) return false;
            uint? valueMetadata = null;
            if (reader.GetAttribute("vm") is string vm) {
                if (!uint.TryParse(vm, out uint parsed)) return false;
                valueMetadata = parsed;
            }

            string? value = null;
            bool hasValue = false;
            using XmlReader child = reader.ReadSubtree();
            while (child.Read()) {
                if (child.NodeType != XmlNodeType.Element || child.Depth != 1) continue;
                if (child.LocalName != "v" || hasValue
                    || !IsSpreadsheetMainNamespace(child.NamespaceURI)) return false;
                hasValue = true;
                value = child.ReadElementContentAsString();
            }
            if (!hasValue && valueMetadata == null) return false;
            snapshot = new DynamicSpillCacheSnapshot {
                Type = parsedType,
                Value = value,
                ValueMetaIndex = valueMetadata
            };
            return true;
        }

        private static bool IsSpreadsheetMainNamespace(string namespaceUri) =>
            string.Equals(namespaceUri, SpreadsheetMainNamespace, StringComparison.Ordinal)
            || string.Equals(namespaceUri, StrictSpreadsheetMainNamespace, StringComparison.Ordinal);
    }
}
