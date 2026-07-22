using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointPicture {
        /// <summary>
        ///     Returns a snapshot of the embedded image bytes for export adapters.
        /// </summary>
        public byte[] GetImageBytes() {
            return GetImageBytes(maximumBytes: null);
        }

        /// <summary>Returns a bounded snapshot of the embedded image bytes.</summary>
        internal byte[] GetImageBytes(int? maximumBytes) {
            ImagePart imagePart = GetImagePart() ?? throw new InvalidOperationException("Picture has no embedded image part.");
            using Stream source = imagePart.GetStream(FileMode.Open, FileAccess.Read);
            return maximumBytes.HasValue
                ? OfficeIMO.Drawing.Internal.OfficeStreamReader.ReadAllBytes(
                    source, maximumBytes.Value)
                : OfficeIMO.Drawing.Internal.OfficeStreamReader.ReadAllBytes(
                    source);
        }
    }
}
