using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointPicture {
        /// <summary>
        ///     Returns a snapshot of the embedded image bytes for export adapters.
        /// </summary>
        public byte[] GetImageBytes() {
            ImagePart imagePart = GetImagePart() ?? throw new InvalidOperationException("Picture has no embedded image part.");
            using Stream source = imagePart.GetStream(FileMode.Open, FileAccess.Read);
            using var buffer = new MemoryStream();
            source.CopyTo(buffer);
            return buffer.ToArray();
        }
    }
}
