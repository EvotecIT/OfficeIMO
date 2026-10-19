using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel.Utilities;
using OfficeIMO.Excel.Xlsb.Model;
using System.Xml;
using System.Xml.Linq;
using System.IO.Compression;

namespace OfficeIMO.Excel.Xlsb.Write;

internal static partial class XlsbNativePackageWriter {
    private static void RebindTableFormulas(ZipArchive archive, XlsbSheetStructure structure, IDictionary<string, byte[]> replacements, ISet<string> deletions,
        int maximumPartBytes, long maximumBytes) {
        if (!structure.NameIndexes.Any(item => item.Value != item.Key)) return;
        using var stream = (FindEntry(archive, "[Content_Types].xml") ?? throw new InvalidDataException("Missing native XLSB content types.")).Open();
        var types = XDocument.Load(stream);
        foreach (var type in types.Root!.Elements().Where(element => (string?)element.Attribute("ContentType") == "application/vnd.ms-excel.table")) {
            string name = ((string?)type.Attribute("PartName") ?? throw new InvalidDataException("A native table part has no package name.")).TrimStart('/');
            if (deletions.Contains(name)) continue;
            byte[] bytes = replacements.TryGetValue(name, out byte[]? replacement) ? replacement
                : ReadEntry(FindEntry(archive, name) ?? throw new InvalidDataException("Missing native XLSB table part."), maximumPartBytes);
            replacements[name] = XlsbFormulaRecordRebinder.RebindPart(bytes, structure.NameIndexes, maximumBytes);
        }
    }

    private static void CopyRenamedRelatedParts(ExcelDocument document, XlsbWorkbook source, IDictionary<string, byte[]> replacements) {
        var renames = document.NativeXlsbPreservation.ReverseRenames(document);
        if (renames.Count == 0) return;
        var parts = new Dictionary<string, OpenXmlPart>(StringComparer.Ordinal);
        var queue = new Queue<OpenXmlPart>(document.OpenXmlDocument.Parts.Select(pair => pair.OpenXmlPart));
        while (queue.Count > 0) {
            var part = queue.Dequeue(); if (parts.ContainsKey(part.Uri.OriginalString)) continue;
            parts.Add(part.Uri.OriginalString, part);
            foreach (var child in part.Parts) queue.Enqueue(child.OpenXmlPart);
        }
        foreach (var mapping in source.ProjectionPartUris) {
            if (!parts.TryGetValue(mapping.Value, out var part) || part.RootElement == null) continue;
            if (ExcelPackageReferenceNormalizer.Copy(part.RootElement, renames).OuterXml == part.RootElement.OuterXml) continue;
            using var output = new MemoryStream();
            using (var writer = XmlWriter.Create(output, new XmlWriterSettings { Encoding = new UTF8Encoding(false), CloseOutput = false })) part.RootElement.WriteTo(writer);
            replacements[mapping.Key] = output.ToArray();
        }
    }
}
