using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.Threading;

namespace OfficeIMO.Visio;

internal static partial class VisioLegacyXmlCodec {
    private static readonly XNamespace CellMetadata = VisioNativeCellMetadata.Namespace;
    private static readonly HashSet<string> ModernErrors = new("#DIM! #DIV/0! #VALUE! #REF! #NUM! #N/A".Split(' '), StringComparer.Ordinal);

    private sealed class LegacyCellState {
        internal LegacyCellState(string? condition, string? error) { Condition = condition; Error = error; }
        internal string? Condition { get; }
        internal string? Error { get; }
    }

    /// <summary>
    /// Stores legacy null-string conditions and unrepresentable errors in the document's permitted foreign-namespace
    /// extension area. Modern Cell@V remains the literal cache; no extra Cell attributes are added.
    /// </summary>
    private static void CaptureLegacyCellMetadata(XElement root, VisioXmlConversionReport report, CancellationToken cancellationToken) {
        var metadata = new XElement(CellMetadata + "NativeCellValues");
        foreach (XElement cell in root.Descendants(Modern + "Cell")) {
            cancellationToken.ThrowIfCancellationRequested();
            if (cell.Annotation<LegacyCellState>() is not LegacyCellState state) continue;
            var entry = new XElement(CellMetadata + "Cell", new XAttribute("Address", CellAddress(cell)), CellValueSnapshot(cell));
            entry.SetAttributeValue("Condition", state.Condition);
            entry.SetAttributeValue("Error", state.Error);
            metadata.Add(entry);
            if (state.Error != null) report.Add("VDX_CELL_ERROR", "Legacy error state has no modern error-token mapping and is preserved as conversion metadata.",
                OfficeConversionLossKind.Approximation, CellAddress(cell));
        }
        if (metadata.HasElements) root.Add(metadata);
    }

    /// <summary>
    /// Restores a native condition only while its value, formula, unit and error state match
    /// the import snapshot. The extension is read as XML nodes, within the shared package
    /// loader's XML byte/character limits; it contains no separately parsed payload or XPath.
    /// </summary>
    private static void RestoreLegacyCellMetadata(XElement root) {
        List<XElement> metadata = root.Elements(CellMetadata + "NativeCellValues").ToList();
        if (metadata.Count == 0) return;
        var addresses = new HashSet<string>(metadata.SelectMany(container => container.Elements(CellMetadata + "Cell"))
            .Attributes("Address").Select(attribute => attribute.Value), StringComparer.Ordinal);
        var cells = new Dictionary<string, XElement?>(StringComparer.Ordinal);
        foreach (XElement cell in root.Descendants(Modern + "Cell")) {
            string address = CellAddress(cell);
            if (!addresses.Contains(address)) continue;
            // A duplicate identity is ambiguous. Do not attach a marker to an arbitrary cell.
            if (cells.ContainsKey(address)) cells[address] = null;
            else cells.Add(address, cell);
        }
        foreach (XElement container in metadata) {
            foreach (XElement entry in container.Elements(CellMetadata + "Cell")) {
                string? address = (string?)entry.Attribute("Address");
                string? condition = (string?)entry.Attribute("Condition");
                string? error = (string?)entry.Attribute("Error");
                if (address == null || (condition == null && error == null) || !cells.TryGetValue(address, out XElement? cell) || cell == null) continue;
                if (XNode.DeepEquals(CellValueSnapshot(cell), entry.Element(CellMetadata + "Snapshot")))
                    cell.AddAnnotation(new LegacyCellState(condition, error));
            }
            // This is package preservation data, not legacy document content.
            container.Remove();
        }
    }

    private static XElement CellValueSnapshot(XElement cell) => VisioNativeCellMetadata.Snapshot(cell);
    private static string CellAddress(XElement cell) => VisioNativeCellMetadata.Address(cell);
}
