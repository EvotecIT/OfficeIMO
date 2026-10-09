using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Visio;

public partial class VisioDocument {
    private void TransferImportedStyleCellMetadata(XElement inserted, IReadOnlyDictionary<string, XElement?> nativeCells, string destinationAddress) {
        var metadata = new XElement(VisioNativeCellMetadata.Namespace + "NativeCellValues");
        string sourceAddress = VisioNativeCellMetadata.Address(inserted);
        foreach (XElement cell in inserted.DescendantsAndSelf(XName.Get("Cell", VisioNamespace))) {
            string address = VisioNativeCellMetadata.Address(cell);
            if (nativeCells.TryGetValue(address, out XElement? entry) && entry != null && VisioNativeCellMetadata.Matches(entry, cell)) {
                // Validate at the source before moving only the address. A retained named
                // row may have a different IX; its added cells follow that destination row.
                var copy = new XElement(entry);
                copy.SetAttributeValue("Address", destinationAddress + address.Substring(sourceAddress.Length));
                metadata.Add(copy);
            }
        }
        if (metadata.HasElements) PreservedDocumentElements.Add(metadata);
    }
}
