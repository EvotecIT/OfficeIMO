using System.Threading;
using System.IO.Packaging;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Visio {
    public partial class VisioDocument {
        /// <summary>Imports a bounded Visio 2003–2010 VSD drawing, VSS stencil or VST template with a loss report.</summary>
        /// <remarks>The binary source is never associated with Save. Save writes a modern package of the imported family.</remarks>
        public static OfficeConversionResult<VisioDocument, OfficeLegacyImportReport> LoadLegacyBinary(string path,
            VisioLegacyBinaryImportOptions? options = null, CancellationToken cancellationToken = default) {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
            VisioPackageType family = Path.GetExtension(path).ToLowerInvariant() switch {
                ".vsd" => VisioPackageType.Drawing, ".vss" => VisioPackageType.Stencil, ".vst" => VisioPackageType.Template,
                _ => throw new NotSupportedException("Legacy binary import supports .vsd, .vss and .vst paths.")
            };
            cancellationToken.ThrowIfCancellationRequested();
            using var stream = File.OpenRead(path);
            return LoadLegacyBinary(stream, family, options, cancellationToken);
        }

        /// <summary>Imports the binary profile from a caller-owned stream, returning the existing editable Visio model.</summary>
        /// <remarks>Seekable streams are read from the beginning and their position is restored. Nonseekable streams are read from their current position. Specify the drawing, stencil or template family explicitly when needed.</remarks>
        public static OfficeConversionResult<VisioDocument, OfficeLegacyImportReport> LoadLegacyBinary(Stream stream,
            VisioPackageType packageType = VisioPackageType.Drawing, VisioLegacyBinaryImportOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (packageType is not (VisioPackageType.Drawing or VisioPackageType.Stencil or VisioPackageType.Template))
                throw new NotSupportedException("Binary import reconstructs macro-free drawing, stencil and template packages.");
            VisioLegacyBinaryImportOptions operation = (options ?? new()).Snapshot();
            byte[] bytes = OfficeStreamReader.ReadAllBytes(stream, cancellationToken, operation.Limits.MaxInputBytes);
            var codec = new VisioLegacyBinaryCodec(bytes, operation, cancellationToken);
            var xml = codec.Read();
            var normalizationReport = new VisioXmlConversionReport();
            using MemoryStream normalized = VisioLegacyXmlCodec.ToPackage(xml, packageType, normalizationReport);
            cancellationToken.ThrowIfCancellationRequested();
            using Package package = Package.Open(normalized, FileMode.Open, FileAccess.Read);
            VisioDocument document = LoadCore(package, filePath: null);
            return new OfficeConversionResult<VisioDocument, OfficeLegacyImportReport>(document, codec.CreateReport(normalizationReport));
        }
    }
}
