using System.IO;

namespace OfficeIMO.Core.Internal {
    /// <summary>Signals that valid compressed data exceeded its configured decoded-size boundary.</summary>
    internal sealed class OfficeDecompressionSizeLimitException : IOException {
        internal OfficeDecompressionSizeLimitException(string message) : base(message) { }
    }
}
