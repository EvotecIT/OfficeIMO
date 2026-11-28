using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Excel.OpenXml;
using OfficeIMO.OpenXml.Internal;
using System.Globalization;

namespace OfficeIMO.Excel {
    // One immutable lookup serves object-model and forward-only cached reads.
    // Unknown rich values retain their ordinary cell fallback and package data.
    internal sealed class RichValueErrorLookup {
        internal static readonly RichValueErrorLookup Empty = new RichValueErrorLookup(new Dictionary<uint, string>());
        private const int Spreadsheet = OfficeXmlNamespaces.SpreadsheetML;
        private const string RichData = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata";
        private const string RichValueExtension = "{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}";
        private const int MaximumEntries = 100_000;
        internal const int MaximumPartBytes = 16 * 1024 * 1024;
        private readonly IReadOnlyDictionary<uint, string> _errors;

        private RichValueErrorLookup(IReadOnlyDictionary<uint, string> errors) { _errors = errors; }
        internal string? Resolve(string? index, string? fallback) =>
            uint.TryParse(index, NumberStyles.None, CultureInfo.InvariantCulture, out uint parsed)
                ? Resolve(parsed, fallback) : fallback;
        internal string? Resolve(uint index, string? fallback) =>
            fallback != null && _errors.TryGetValue(index, out string? error) ? error : fallback;

        internal static RichValueErrorLookup FromWorkbook(object? workbook, int maximumPartBytes = MaximumPartBytes) =>
            FromWorkbook(ExcelPartRef.FromSdk(workbook), maximumPartBytes);

        internal static RichValueErrorLookup FromWorkbook(ExcelPartRef workbook, int maximumPartBytes) {
            ExcelPartRef metadata = workbook.FindWorkbookPart(ExcelReadWorkbookParts.Metadata);
            ExcelPartRef values = workbook.FindWorkbookPart(ExcelReadWorkbookParts.RichValues);
            ExcelPartRef structures = workbook.FindWorkbookPart(ExcelReadWorkbookParts.RichStructures);
            if (metadata.IsNone || values.IsNone || structures.IsNone) return Empty;
            foreach (ExcelPartRef part in new[] { metadata, values, structures }) {
                using Stream stream = part.OpenRead();
                CheckPartLength(stream.Length, maximumPartBytes);
            }
            return FromRoots(metadata.Root, values.Root, structures.Root);
        }

        internal static void CheckPartLength(long length, int maximumPartBytes) {
            if (length > maximumPartBytes) throw new InvalidDataException($"Rich-value metadata exceeds its {maximumPartBytes}-byte part limit.");
        }

        internal static RichValueErrorLookup FromRoots(object? metadata, object? values, object? structures) =>
            FromRoots(OfficeXmlNode.FromSdk(metadata), OfficeXmlNode.FromSdk(values), OfficeXmlNode.FromSdk(structures));

        internal static RichValueErrorLookup FromRoots(OfficeXmlNode metadata, OfficeXmlNode values, OfficeXmlNode structures) {
            if (!metadata.Is(Spreadsheet, "metadata") || !values.Is(RichData, "rvData") || !structures.Is(RichData, "rvStructures")) return Empty;
            OfficeXmlNode[] types = Bounded(metadata.Child(Spreadsheet, "metadataTypes"), "metadataType");
            OfficeXmlNode[] blocks = Bounded(metadata.Child(Spreadsheet, "valueMetadata"), "bk");
            OfficeXmlNode richMetadata = default;
            foreach (OfficeXmlNode futureMetadata in metadata.Children(Spreadsheet, "futureMetadata")) {
                if (futureMetadata.Attribute("name") == "XLRICHVALUE") {
                    richMetadata = futureMetadata;
                    break;
                }
            }
            OfficeXmlNode[] future = Bounded(richMetadata, "bk");
            OfficeXmlNode[] richValues = Bounded(values, "rv", rich: true);
            OfficeXmlNode[] richStructures = Bounded(structures, "s", rich: true);
            var errors = new Dictionary<uint, string>();
            // References may alias the same large payload. Cache successful and failed
            // decodes for this operation only, so editable roots remain observable later.
            var futureErrors = new Dictionary<uint, string?>();
            var valueErrors = new Dictionary<uint, string?>();
            var keyLayouts = new Dictionary<uint, (int Count, int ErrorIndex)>();
            string? DecodeValue(uint valueIndex) {
                if (valueErrors.TryGetValue(valueIndex, out string? cached)) return cached;
                string? error = null;
                OfficeXmlNode value = richValues[valueIndex];
                if (TryIndex(value, "s", out uint structureIndex) && structureIndex < richStructures.Length) {
                    OfficeXmlNode structure = richStructures[structureIndex];
                    if (structure.Attribute("t") == "_error") {
                        if (!keyLayouts.TryGetValue(structureIndex, out var layout)) {
                            OfficeXmlNode[] keys = Bounded(structure, "k", rich: true);
                            int errorIndex = Array.FindIndex(keys, key => key.Attribute("n") == "errorType" && key.Attribute("t") == "i");
                            if (keys.Count(key => key.Attribute("n") == "errorType") != 1) errorIndex = -1;
                            layout = (keys.Length, errorIndex);
                            keyLayouts.Add(structureIndex, layout);
                        }
                        OfficeXmlNode[] entries = Bounded(value, "v", rich: true);
                        if (layout.Count == entries.Length && layout.ErrorIndex >= 0) {
                            string code = entries[layout.ErrorIndex].InnerText;
                            error = code == "8" ? "#SPILL!" : code == "13" ? "#CALC!" : null;
                        }
                    }
                }
                valueErrors.Add(valueIndex, error);
                return error;
            }
            string? DecodeFuture(uint futureIndex) {
                if (futureErrors.TryGetValue(futureIndex, out string? cached)) return cached;
                OfficeXmlNode valueBlock = default;
                foreach (OfficeXmlNode extension in future[futureIndex].Child(Spreadsheet, "extLst").Children(Spreadsheet, "ext")) {
                    if (!string.Equals(extension.Attribute("uri"), RichValueExtension, StringComparison.OrdinalIgnoreCase)) continue;
                    valueBlock = extension.Child(RichData, "rvb");
                    if (valueBlock.Exists) break;
                }
                string? error = valueBlock.Exists && TryIndex(valueBlock, "i", out uint valueIndex)
                    && valueIndex < richValues.Length ? DecodeValue(valueIndex) : null;
                futureErrors.Add(futureIndex, error);
                return error;
            }
            for (int i = 0; i < blocks.Length; i++) {
                foreach (OfficeXmlNode record in blocks[i].Children(Spreadsheet, "rc")) {
                    if (!TryIndex(record, "t", out uint type) || type == 0 || type > types.Length
                        || types[type - 1].Attribute("name") != "XLRICHVALUE"
                        || !TryIndex(record, "v", out uint futureIndex) || futureIndex >= future.Length) continue;
                    string? error = DecodeFuture(futureIndex);
                    if (error != null) errors[(uint)i + 1] = error;
                }
            }
            return new RichValueErrorLookup(errors);
        }

        private static bool TryIndex(OfficeXmlNode element, string name, out uint value) =>
            uint.TryParse(element.Attribute(name), NumberStyles.None, CultureInfo.InvariantCulture, out value);

        private static OfficeXmlNode[] Bounded(OfficeXmlNode owner, string name, bool rich = false) {
            var children = new List<OfficeXmlNode>();
            foreach (OfficeXmlNode child in owner.Children()) {
                if (!(rich ? child.Is(RichData, name) : child.Is(Spreadsheet, name))) continue;
                if (children.Count == MaximumEntries) throw new InvalidDataException("Rich-value metadata exceeds its 100,000-entry limit.");
                children.Add(child);
            }
            return children.ToArray();
        }
    }
}
