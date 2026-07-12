namespace OfficeIMO.Word {
    public partial class WordDocument {
        /// <summary>
        /// Returns the bytes of all embedded images in the document.
        /// Images linked externally are skipped.
        /// </summary>
        public IReadOnlyList<byte[]> GetImageBytes() {
            List<byte[]> images = new List<byte[]>();
            foreach (var img in Images) {
                try {
                    images.Add(img.ToBytes());
                } catch (InvalidOperationException) {
                    // external image - skip
                }
            }
            return images;
        }

        /// <summary>
        /// Returns streams with data of all embedded images in the document.
        /// Images linked externally are skipped.
        /// </summary>
        public IReadOnlyList<Stream> OpenImageStreams() {
            List<Stream> streams = new List<Stream>();
            foreach (var img in Images) {
                try {
                    streams.Add(img.OpenRead());
                } catch (InvalidOperationException) {
                    // external image - skip
                }
            }
            return streams;
        }
    }
}
