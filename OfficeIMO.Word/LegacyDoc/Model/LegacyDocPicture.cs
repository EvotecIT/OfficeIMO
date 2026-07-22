namespace OfficeIMO.Word.LegacyDoc.Model {
    internal sealed class LegacyDocPicture {
        private readonly byte[] _imageBytes;

        internal LegacyDocPicture(byte[] imageBytes, string contentType, double widthPixels, double heightPixels) {
            _imageBytes = imageBytes == null ? throw new ArgumentNullException(nameof(imageBytes)) : (byte[])imageBytes.Clone();
            ContentType = contentType ?? throw new ArgumentNullException(nameof(contentType));
            WidthPixels = widthPixels;
            HeightPixels = heightPixels;
        }

        internal byte[] ImageBytes => (byte[])_imageBytes.Clone();

        internal int ImageByteCount => _imageBytes.Length;

        internal string ContentType { get; }

        internal double WidthPixels { get; }

        internal double HeightPixels { get; }

        internal string FileName => ContentType.ToLowerInvariant() switch {
            "image/png" => "legacy-picture.png",
            "image/jpeg" => "legacy-picture.jpg",
            "image/bmp" => "legacy-picture.bmp",
            "image/tiff" => "legacy-picture.tiff",
            "image/x-emf" => "legacy-picture.emf",
            "image/x-wmf" => "legacy-picture.wmf",
            _ => "legacy-picture.bin"
        };
    }
}
