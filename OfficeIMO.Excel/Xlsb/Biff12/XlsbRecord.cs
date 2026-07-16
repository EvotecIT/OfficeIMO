namespace OfficeIMO.Excel.Xlsb.Biff12 {
    /// <summary>
    /// Represents one framed BIFF12 record from an XLSB binary part.
    /// </summary>
    internal sealed class XlsbRecord {
        internal XlsbRecord(long offset, int headerSize, int type, byte[] data) {
            Offset = offset;
            HeaderSize = headerSize;
            Type = type;
            Data = data ?? throw new ArgumentNullException(nameof(data));
        }

        /// <summary>Gets the zero-based stream offset of the record header, or -1 for a non-seekable stream.</summary>
        internal long Offset { get; }

        /// <summary>Gets the encoded header length in bytes.</summary>
        internal int HeaderSize { get; }

        /// <summary>Gets the BIFF12 record type number.</summary>
        internal int Type { get; }

        /// <summary>Gets the record-specific payload.</summary>
        internal byte[] Data { get; }

        /// <summary>Gets the payload size declared by the record header.</summary>
        internal int Size => Data.Length;
    }
}
