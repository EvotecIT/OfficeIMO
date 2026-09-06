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
        Guard.NotNull(stream, nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("The invoice XML stream must be readable.", nameof(stream));
        using (var buffer = new MemoryStream()) {
            byte[] chunk = new byte[8192];
            while (true) {
                int count = stream.Read(chunk, 0, (int)Math.Min(chunk.Length, MaximumXmlBytes + 1L - buffer.Length));
                if (count == 0) break;
                buffer.Write(chunk, 0, count);
                if (buffer.Length > MaximumXmlBytes) throw new InvalidDataException("CII XML exceeds the maximum byte length.");
            }
            return Load(buffer.ToArray());
        }
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
        File.WriteAllBytes(path, _bytes);
    }
}
