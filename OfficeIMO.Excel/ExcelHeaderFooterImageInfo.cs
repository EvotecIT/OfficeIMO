namespace OfficeIMO.Excel {
    /// <summary>
    /// Snapshot of an Excel header/footer image.
    /// </summary>
    public sealed class ExcelHeaderFooterImageInfo {
        private readonly byte[] _bytes;

        internal ExcelHeaderFooterImageInfo(ExcelHeaderFooterPosition position, byte[] bytes, string contentType, OfficeLength width, OfficeLength height) {
            Position = position;
            _bytes = (byte[])bytes.Clone();
            ContentType = contentType;
            Width = width;
            Height = height;
        }

        /// <summary>Header/footer section position.</summary>
        public ExcelHeaderFooterPosition Position { get; }
        /// <summary>Image bytes.</summary>
        public byte[] Bytes => (byte[])_bytes.Clone();
        /// <summary>Image content type, such as image/png or image/jpeg.</summary>
        public string ContentType { get; }
        /// <summary>Image physical width.</summary>
        public OfficeLength Width { get; }
        /// <summary>Image physical height.</summary>
        public OfficeLength Height { get; }
        internal long SourceByteCount => _bytes.LongLength;
    }
}
