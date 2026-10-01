namespace OfficeIMO.Pdf;

public sealed partial class PdfCiiInvoiceDocument {
    /// <summary>Loads existing CII invoice XML from a file, applying the same limits as byte-array loading.</summary>
    public static PdfCiiInvoiceDocument Load(string path) {
        Guard.NotNull(path, nameof(path));
        using (var stream = File.OpenRead(path)) {
            return Load(stream);
        }
    }

    /// <summary>
    /// Loads existing CII invoice XML from the stream's current position to its end.
    /// Reads at most the size limit plus one byte and leaves the caller's stream open.
    /// </summary>
    public static PdfCiiInvoiceDocument Load(Stream stream) {
        return new PdfCiiInvoiceDocument(OfficeIMO.Internal.Invoicing.InvoicePreservedXml.Load(stream));
    }

    /// <summary>Writes the stored XML bytes at the current position and leaves the caller's stream open.</summary>
    public void Save(Stream stream) {
        Guard.NotNull(stream, nameof(stream));
        byte[] snapshot = ToBytes();
        stream.Write(snapshot, 0, snapshot.Length);
    }

    /// <summary>Saves the stored XML bytes to a file, replacing any existing contents.</summary>
    public void Save(string path) {
        Guard.NotNull(path, nameof(path));
        File.WriteAllBytes(path, ToBytes());
    }
}
