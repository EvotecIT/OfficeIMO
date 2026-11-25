using System.Security.Cryptography;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        private Dictionary<string, ReusableImagePart>? _imagePartsByContent;

        /// <summary>
        ///     Returns an image part this document already created with the same content, so a picture added many
        ///     times is stored once, as PowerPoint does. Parts deleted since then, or moved to another package, are ignored.
        /// </summary>
        internal bool TryGetReusableImagePart(string key, OpenXmlPackage package, out ReusableImagePart image) {
            if (_imagePartsByContent != null && _imagePartsByContent.TryGetValue(key, out ReusableImagePart? cached)) {
                if (IsLivePart(cached.Part, package)) {
                    image = cached;
                    return true;
                }
                _imagePartsByContent.Remove(key);
            }
            image = null!;
            return false;
        }

        internal void RegisterReusableImagePart(string key, ReusableImagePart image) {
            (_imagePartsByContent ??= new Dictionary<string, ReusableImagePart>(StringComparer.Ordinal))[key] = image;
        }

        internal static string CreateImageContentKey(string kind, byte[] data) {
#if NET5_0_OR_GREATER
            byte[] hash = SHA256.HashData(data);
#else
            byte[] hash;
            using (SHA256 sha = SHA256.Create()) hash = sha.ComputeHash(data);
#endif
            return kind + ":" + data.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + ":" + System.Convert.ToBase64String(hash);
        }

        private static bool IsLivePart(OpenXmlPart part, OpenXmlPackage package) {
            try {
                // A deleted part no longer has a URI.
                _ = part.Uri;
                return ReferenceEquals(part.OpenXmlPackage, package);
            } catch (InvalidOperationException) {
                return false;
            }
        }
    }

    /// <summary>An image part with the pixel size of the raster stored in it, when known.</summary>
    internal sealed class ReusableImagePart {
        internal ReusableImagePart(ImagePart part, int width, int height) {
            Part = part;
            Width = width;
            Height = height;
        }

        internal ImagePart Part { get; }

        internal int Width { get; }

        internal int Height { get; }
    }
}
