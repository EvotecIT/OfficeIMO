using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Utilities;
using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Model;
using System.Xml.Linq;

namespace OfficeIMO.Excel.Xlsb.Write;

/// <summary>Rebinds native tab identities and package relationships while retaining the original binary record graph.</summary>
internal sealed class XlsbSheetStructure {
    private static readonly XNamespace Relationships = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static readonly XNamespace ContentTypes = "http://schemas.openxmlformats.org/package/2006/content-types";
    private const string WorksheetType = "application/vnd.ms-excel.worksheet";
    private readonly XlsbWorkbook _source;
    internal IReadOnlyList<ExcelLegacySheetBindings.Binding> Bindings { get; }
    internal Dictionary<string, string> PartNames { get; } = new(StringComparer.Ordinal);
    internal Dictionary<string, string> NewCodeNames { get; } = new(StringComparer.Ordinal);
    internal bool Changed { get; }
    internal Dictionary<uint, uint?> NameIndexes { get; } = new();

    internal XlsbSheetStructure(ExcelDocument document, XlsbWorkbook source) {
        _source = source; Bindings = document.NativeXlsbPreservation.ReadCurrent(document);
        if (Bindings.Count == 0 || Bindings.Count > 32767 || Bindings.Select(binding => binding.Name).Distinct(StringComparer.OrdinalIgnoreCase).Count() != Bindings.Count)
            throw new NotSupportedException("Native XLSB requires a nonempty collection of uniquely named sheets.");
        Changed = Bindings.Count != source.Sheets.Count || Bindings.Any(binding => binding.Original == null || binding.Original.NativeIndex != binding.Index
            || binding.Name != binding.Original.Name || State(binding) != source.Sheets[binding.Original.NativeIndex].State);
        uint keptNames = 0;
        for (int index = 0; index < source.DefinedNames.Count; index++) {
            uint scope = source.DefinedNames[index].LocalSheetIndex;
            NameIndexes[checked((uint)index + 1)] = scope != uint.MaxValue && Map(checked((int)scope)) < 0 ? null : ++keptNames;
        }
        foreach (var binding in Bindings.Where(binding => binding.Original != null))
            PartNames.Add(binding.Part.Uri.OriginalString, source.Sheets[binding.Original!.NativeIndex].PartName!);
    }

    internal int Map(int oldIndex) => Bindings.FirstOrDefault(binding => binding.Original?.NativeIndex == oldIndex)?.Index ?? -1;
    internal static uint State(ExcelLegacySheetBindings.Binding binding) => binding.Element.State?.Value == SheetStateValues.VeryHidden ? 2U
        : binding.Element.State?.Value == SheetStateValues.Hidden ? 1U : 0U;

    internal void PreparePackage(ExcelDocument document, string workbookPartName, IReadOnlyCollection<string> existingNames,
        byte[] workbookBytes, byte[] relationshipBytes, byte[] contentTypeBytes, IDictionary<string, byte[]> replacements, ISet<string> deletions, long maximumBytes) {
        if (!Changed) return;
        var relationships = XDocument.Load(new MemoryStream(relationshipBytes));
        var types = XDocument.Load(new MemoryStream(contentTypeBytes));
        var ids = new HashSet<string>(relationships.Root!.Elements().Select(element => (string?)element.Attribute("Id") ?? string.Empty), StringComparer.Ordinal);
        var paths = new HashSet<string>(existingNames, StringComparer.OrdinalIgnoreCase);
        uint tabId = _source.Sheets.Select(sheet => sheet.TabId).DefaultIfEmpty().Max();
        string directory = workbookPartName.Substring(0, Math.Max(0, workbookPartName.LastIndexOf('/') + 1));
        var bundles = new List<byte[]>();
        foreach (var binding in Bindings) {
            string relationshipId; uint nativeId;
            if (binding.Original != null) {
                var original = _source.Sheets[binding.Original.NativeIndex]; relationshipId = original.RelationshipId; nativeId = original.TabId;
            } else {
                int index = 1; string path;
                do { path = directory + "worksheets/sheet" + index++ + ".bin"; } while (!paths.Add(path));
                PartNames.Add(binding.Part.Uri.OriginalString, path);
                index = 1; do { relationshipId = "rId" + index++; } while (!ids.Add(relationshipId));
                nativeId = checked(++tabId);
                relationships.Root!.Add(new XElement(Relationships + "Relationship", new XAttribute("Id", relationshipId),
                    new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"),
                    new XAttribute("Target", path.Substring(directory.Length))));
                types.Root!.Add(new XElement(ContentTypes + "Override", new XAttribute("PartName", "/" + path), new XAttribute("ContentType", WorksheetType)));
            }
            using var payload = new MemoryStream();
            Put(payload, State(binding)); Put(payload, nativeId); Wide(payload, relationshipId); Wide(payload, binding.Name);
            bundles.Add(payload.ToArray());
        }
        foreach (var old in _source.Sheets.Select((sheet, index) => (sheet, index)).Where(item => Map(item.index) < 0)) {
            relationships.Root!.Elements().Where(element => (string?)element.Attribute("Id") == old.sheet.RelationshipId).Remove();
            if (old.sheet.PartName != null) {
                // Unreferenced related parts remain valid package carriers; never delete a shared chart/image graph.
                deletions.Add(old.sheet.PartName); deletions.Add(RelationshipPartName(old.sheet.PartName));
                types.Root!.Elements().Where(element => (string?)element.Attribute("PartName") == "/" + old.sheet.PartName).Remove();
            }
        }
        replacements[workbookPartName] = RewriteWorkbook(workbookBytes, bundles, maximumBytes);
        replacements[RelationshipPartName(workbookPartName)] = XmlBytes(relationships);
        replacements["[Content_Types].xml"] = XmlBytes(types);
        if (_source.VbaProjectPartName != null) {
            byte[] bytes = _source.RelatedParts[_source.VbaProjectPartName].Bytes;
            byte[] edited = ExcelLegacyVbaHostEditor.Reconcile(bytes, Bindings,
                _source.Worksheets.ToDictionary(sheet => sheet.Name, sheet => (string?)sheet.Properties?.CodeName, StringComparer.Ordinal),
                document.NativeXlsbPreservation.AllocatedCodeNames, NewCodeNames);
            if (!ReferenceEquals(bytes, edited)) replacements[_source.VbaProjectPartName] = edited;
        }
    }

    private byte[] RewriteWorkbook(byte[] original, IReadOnlyList<byte[]> bundles, long maximumBytes) {
        using var input = new MemoryStream(original); var records = XlsbRecordReader.ReadAll(input);
        using var output = new OfficeIMO.Core.Internal.OfficeBoundedMemoryStream(maximumBytes);
        bool written = false;
        foreach (var record in records) {
            if (record.Type == 156) {
                if (!written) foreach (byte[] bundle in bundles) XlsbRecordWriter.Write(output, 156, bundle);
                written = true; continue;
            }
            byte[] payload = record.Data;
            if (record.Type == 362) {
                payload = (byte[])payload.Clone();
                int count = checked((int)Read(payload, 0));
                if (payload.Length != 4 + checked(count * 12)) throw new InvalidDataException("Invalid XLSB external-sheet table.");
                for (int index = 0; index < count; index++) {
                    int offset = 4 + index * 12; uint book = Read(payload, offset);
                    if (book >= _source.SelfSupportingLinks.Count || !_source.SelfSupportingLinks[(int)book]) continue;
                    int first = unchecked((int)Read(payload, offset + 4)), last = unchecked((int)Read(payload, offset + 8));
                    if (first < 0 || last < 0) continue;
                    if (first >= _source.Sheets.Count || last >= _source.Sheets.Count) throw new InvalidDataException("The XLSB self reference points beyond its sheet collection.");
                    var survivors = Enumerable.Range(Math.Min(first, last), checked(Math.Abs(last - first) + 1)).Where(old => Map(old) >= 0).ToArray();
                    int start = survivors.Length == 0 ? -1 : Map(first <= last ? survivors[0] : survivors[survivors.Length - 1]);
                    int end = survivors.Length == 0 ? -1 : Map(first <= last ? survivors[survivors.Length - 1] : survivors[0]);
                    Put(payload, offset + 4, unchecked((uint)start)); Put(payload, offset + 8, unchecked((uint)end));
                }
            } else if (record.Type == 39) {
                payload = (byte[])payload.Clone(); uint scope = Read(payload, 5);
                if (scope != uint.MaxValue) {
                    int mapped = Map(checked((int)scope));
                    if (mapped < 0) continue;
                    Put(payload, 5, checked((uint)mapped));
                }
            } else if (record.Type == 158 && payload.Length >= 29) {
                payload = (byte[])payload.Clone();
                foreach (int offset in new[] { 20, 24 }) Put(payload, offset, checked((uint)Math.Max(0, Map(checked((int)Read(payload, offset))))));
            }
            if (record.Type == 39 && NameIndexes.Any(item => item.Value != item.Key)) payload = XlsbFormulaRecordRebinder.Rebind(record.Type, payload, NameIndexes);
            XlsbRecordWriter.Write(output, record.Type, payload);
        }
        if (!written) throw new InvalidDataException("The source XLSB workbook has no sheet bundle records.");
        return output.ToArray();
    }

    internal static string RelationshipPartName(string path) {
        int separator = path.LastIndexOf('/'); return path.Substring(0, separator + 1) + "_rels/" + path.Substring(separator + 1) + ".rels";
    }
    private static byte[] XmlBytes(XDocument document) { using var stream = new MemoryStream(); document.Save(stream); return stream.ToArray(); }
    private static uint Read(byte[] payload, int offset) {
        if (offset < 0 || offset > payload.Length - 4) throw new InvalidDataException("A native XLSB binding is truncated.");
        return (uint)(payload[offset] | payload[offset + 1] << 8 | payload[offset + 2] << 16 | payload[offset + 3] << 24);
    }
    private static void Put(byte[] payload, int offset, uint value) { for (int index = 0; index < 4; index++) payload[offset + index] = (byte)(value >> (index * 8)); }
    private static void Put(Stream output, uint value) { for (int index = 0; index < 4; index++) output.WriteByte((byte)(value >> (index * 8))); }
    private static void Wide(Stream output, string value) { Put(output, checked((uint)value.Length)); byte[] bytes = Encoding.Unicode.GetBytes(value); output.Write(bytes, 0, bytes.Length); }
}
