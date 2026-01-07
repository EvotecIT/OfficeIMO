using System;
using System.IO;
using System.IO.Packaging;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Load orchestrator for VisioDocument.
    /// </summary>
    public partial class VisioDocument {
        /// <summary>
        /// Loads an existing .vsdx file into a VisioDocument.
        /// </summary>
        public static VisioDocument Load(string filePath) => LoadCore(filePath);

        /// <summary>
        /// Loads an existing .vsdx document from a stream.
        /// </summary>
        public static VisioDocument Load(Stream stream) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

            if (stream.CanSeek) {
                stream.Seek(0, SeekOrigin.Begin);
            }

            using var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            buffer.Seek(0, SeekOrigin.Begin);

            using Package package = Package.Open(buffer, FileMode.Open, FileAccess.Read);
            return LoadCore(package, filePath: null);
        }
    }
}

