using System;
using System.IO;
using System.Linq;

namespace OfficeIMO.Visio.Stencils {
    /// <summary>
    /// Embedded preview/icon image payload extracted from a package-backed stencil master.
    /// </summary>
    public sealed class VisioStencilPreviewImageData {
        private readonly byte[] _data;

        /// <summary>
        /// Initializes extracted preview image payload data.
        /// </summary>
        public VisioStencilPreviewImageData(string masterId, string masterNameU, string? masterName, VisioStencilPreviewImage previewImage, byte[] data) {
            if (string.IsNullOrWhiteSpace(masterId)) throw new ArgumentException("Master id cannot be null or whitespace.", nameof(masterId));
            if (string.IsNullOrWhiteSpace(masterNameU)) throw new ArgumentException("Master NameU cannot be null or whitespace.", nameof(masterNameU));
            if (previewImage == null) throw new ArgumentNullException(nameof(previewImage));
            if (data == null) throw new ArgumentNullException(nameof(data));
            if (data.Length == 0) throw new ArgumentException("Preview image payload cannot be empty.", nameof(data));

            MasterId = masterId;
            MasterNameU = masterNameU;
            MasterName = string.IsNullOrWhiteSpace(masterName) ? null : masterName;
            PreviewImage = previewImage;
            _data = data.ToArray();
        }

        /// <summary>Source master identifier.</summary>
        public string MasterId { get; }

        /// <summary>Source master universal name.</summary>
        public string MasterNameU { get; }

        /// <summary>Source master display name, when present.</summary>
        public string? MasterName { get; }

        /// <summary>Preview image metadata associated with the extracted payload.</summary>
        public VisioStencilPreviewImage PreviewImage { get; }

        /// <summary>Embedded preview image byte length.</summary>
        public int ByteLength => _data.Length;

        /// <summary>Copy of the embedded preview image bytes.</summary>
        public byte[] Data => _data.ToArray();

        /// <summary>Stable file name suitable for saving the extracted preview payload.</summary>
        public string SuggestedFileName {
            get {
                string extension = string.IsNullOrWhiteSpace(PreviewImage.Extension)
                    ? ".bin"
                    : PreviewImage.Extension!.StartsWith(".", StringComparison.Ordinal) ? PreviewImage.Extension! : "." + PreviewImage.Extension;
                return Sanitize(MasterId + "-" + MasterNameU) + extension;
            }
        }

        /// <summary>
        /// Saves the preview payload to a file.
        /// </summary>
        public void Save(string path) {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("Output path cannot be null or whitespace.", nameof(path));

            string? directory = Path.GetDirectoryName(Path.GetFullPath(path));
            if (!string.IsNullOrWhiteSpace(directory)) {
                Directory.CreateDirectory(directory!);
            }

            OfficeIMO.Core.Internal.OfficeFileCommit.WriteAllBytes(path, _data);
        }

        /// <summary>
        /// Saves the preview payload to a directory using <see cref="SuggestedFileName"/>.
        /// </summary>
        public string SaveToDirectory(string directoryPath) {
            if (string.IsNullOrWhiteSpace(directoryPath)) throw new ArgumentException("Output directory cannot be null or whitespace.", nameof(directoryPath));

            Directory.CreateDirectory(directoryPath);
            string path = Path.Combine(directoryPath, SuggestedFileName);
            Save(path);
            return path;
        }

        private static string Sanitize(string value) {
            char[] invalid = Path.GetInvalidFileNameChars();
            char[] chars = value
                .Select(ch => invalid.Contains(ch) || char.IsWhiteSpace(ch) ? '-' : ch)
                .ToArray();
            string sanitized = new string(chars).Trim('-');
            return string.IsNullOrWhiteSpace(sanitized) ? "preview" : sanitized;
        }
    }
}
