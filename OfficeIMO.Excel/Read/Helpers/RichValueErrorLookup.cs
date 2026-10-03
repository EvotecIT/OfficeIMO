using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.Xml;
using Rich = DocumentFormat.OpenXml.Office2019.Excel.RichData;

namespace OfficeIMO.Excel {
    // One immutable lookup serves object-model and forward-only cached reads.
    // Unknown rich values retain their ordinary cell fallback and package data.
    internal sealed class RichValueErrorLookup {
        internal static readonly RichValueErrorLookup Empty = new RichValueErrorLookup(new Dictionary<uint, string>());
        private const int MaximumEntries = 100_000;
        internal const int MaximumPartBytes = 16 * 1024 * 1024;
        private readonly IReadOnlyDictionary<uint, string> _errors;

        private RichValueErrorLookup(IReadOnlyDictionary<uint, string> errors) { _errors = errors; }
        internal string? Resolve(string? index, string? fallback) =>
            uint.TryParse(index, NumberStyles.None, CultureInfo.InvariantCulture, out uint parsed)
                ? Resolve(parsed, fallback) : fallback;
        internal string? Resolve(uint index, string? fallback) =>
            fallback != null && _errors.TryGetValue(index, out string? error) ? error : fallback;

        internal static RichValueErrorLookup FromWorkbook(WorkbookPart? workbook, int maximumPartBytes = MaximumPartBytes) {
            if (workbook?.CellMetadataPart == null) return Empty;
            RdRichValuePart? values = workbook.RdRichValueParts.FirstOrDefault();
            RdRichValueStructurePart? structures = workbook.GetPartsOfType<RdRichValueStructurePart>().FirstOrDefault();
            if (values == null || structures == null) return Empty;
            foreach (OpenXmlPart part in new OpenXmlPart[] { workbook.CellMetadataPart, values, structures }) {
                using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
                if (stream.Length > maximumPartBytes) throw new InvalidDataException($"Rich-value metadata exceeds its {maximumPartBytes}-byte part limit.");
            }
            return FromRoots(workbook.CellMetadataPart.Metadata, values.RichValueData, structures.RichValueStructures);
        }

        internal static string ReadXml(Stream stream) {
            using XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null,
                MaxCharactersInDocument = MaximumPartBytes, CloseInput = false
            });
            reader.MoveToContent();
            string xml = reader.ReadOuterXml();
            while (!reader.EOF) reader.Read();
            return xml;
        }

        internal static RichValueErrorLookup FromRoots(Metadata? metadata, Rich.RichValueData? values, Rich.RichValueStructures? structures) {
            if (metadata == null || values == null || structures == null) return Empty;
            MetadataType[] types = Bounded(metadata.MetadataTypes?.Elements<MetadataType>());
            MetadataBlock[] blocks = Bounded(metadata.GetFirstChild<ValueMetadata>()?.Elements<MetadataBlock>());
            FutureMetadataBlock[] future = Bounded(metadata.Elements<FutureMetadata>()
                .FirstOrDefault(item => item.Name?.Value == "XLRICHVALUE")?.Elements<FutureMetadataBlock>());
            Rich.RichValue[] richValues = Bounded(values.Elements<Rich.RichValue>());
            Rich.RichValueStructure[] richStructures = Bounded(structures.Elements<Rich.RichValueStructure>());
            var errors = new Dictionary<uint, string>();
            // References may alias the same large payload. Cache successful and failed
            // decodes for this operation only, so editable roots remain observable later.
            var futureErrors = new Dictionary<uint, string?>();
            var valueErrors = new Dictionary<uint, string?>();
            var keyLayouts = new Dictionary<uint, (int Count, int ErrorIndex)>();
            string? DecodeValue(uint valueIndex) {
                if (valueErrors.TryGetValue(valueIndex, out string? cached)) return cached;
                string? error = null;
                Rich.RichValue value = richValues[valueIndex];
                if (TryIndex(value, "s", out uint structureIndex) && structureIndex < richStructures.Length) {
                    Rich.RichValueStructure structure = richStructures[structureIndex];
                    if (structure.T?.Value == "_error") {
                        if (!keyLayouts.TryGetValue(structureIndex, out var layout)) {
                            Rich.Key[] keys = Bounded(structure.Elements<Rich.Key>());
                            int errorIndex = Array.FindIndex(keys, key => key.N?.Value == "errorType"
                                && key.GetAttributes().Any(attribute => attribute.LocalName == "t"
                                    && string.IsNullOrEmpty(attribute.NamespaceUri) && attribute.Value == "i"));
                            if (keys.Count(key => key.N?.Value == "errorType") != 1) errorIndex = -1;
                            layout = (keys.Length, errorIndex);
                            keyLayouts.Add(structureIndex, layout);
                        }
                        Rich.Value[] entries = Bounded(value.Elements<Rich.Value>());
                        if (layout.Count == entries.Length && layout.ErrorIndex >= 0) {
                            string code = entries[layout.ErrorIndex].Text;
                            error = code == "8" ? "#SPILL!" : code == "13" ? "#CALC!" : null;
                        }
                    }
                }
                valueErrors.Add(valueIndex, error);
                return error;
            }
            string? DecodeFuture(uint futureIndex) {
                if (futureErrors.TryGetValue(futureIndex, out string? cached)) return cached;
                OpenXmlElement? valueBlock = future[futureIndex].Descendants()
                    .FirstOrDefault(item => item.LocalName == "rvb"
                        && item.NamespaceUri == "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata");
                string? error = valueBlock != null && TryIndex(valueBlock, "i", out uint valueIndex)
                    && valueIndex < richValues.Length ? DecodeValue(valueIndex) : null;
                futureErrors.Add(futureIndex, error);
                return error;
            }
            for (int i = 0; i < blocks.Length; i++) {
                foreach (MetadataRecord record in blocks[i].Elements<MetadataRecord>()) {
                    if (!TryIndex(record, "t", out uint type) || type == 0 || type > types.Length
                        || types[type - 1].Name?.Value != "XLRICHVALUE"
                        || !TryIndex(record, "v", out uint futureIndex) || futureIndex >= future.Length) continue;
                    string? error = DecodeFuture(futureIndex);
                    if (error != null) errors[(uint)i + 1] = error;
                }
            }
            return new RichValueErrorLookup(errors);
        }

        private static bool TryIndex(OpenXmlElement element, string name, out uint value) =>
            uint.TryParse(element.GetAttributes().FirstOrDefault(attribute => attribute.LocalName == name && string.IsNullOrEmpty(attribute.NamespaceUri)).Value,
                NumberStyles.None, CultureInfo.InvariantCulture, out value);
        private static T[] Bounded<T>(IEnumerable<T>? elements) {
            T[] result = elements?.Take(MaximumEntries + 1).ToArray() ?? Array.Empty<T>();
            if (result.Length > MaximumEntries) throw new InvalidDataException("Rich-value metadata exceeds its 100,000-entry limit.");
            return result;
        }
    }
}
