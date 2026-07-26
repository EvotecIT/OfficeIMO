using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Applies bounded nested-package validation before embedded chart workbooks are opened.
    /// </summary>
    internal static class PowerPointChartWorkbookSecurity {
        private const int MaximumPackageBytes = 8 * 1024 * 1024;
        private const long MaximumCharactersInPart = 2L * 1024L * 1024L;

        internal static byte[] ReadAndValidate(Stream stream) {
            byte[] workbookBytes = ReadBounded(stream);
            OfficePackageSecurityInspector.Validate(
                workbookBytes,
                CreateSecurityOptions());
            return workbookBytes;
        }

        internal static OpenSettings CreateOpenSettings() => new OpenSettings {
            AutoSave = false,
            MaxCharactersInPart = MaximumCharactersInPart
        };

        private static OfficePackageSecurityOptions CreateSecurityOptions() =>
            new OfficePackageSecurityOptions {
                MaxPackageBytes = MaximumPackageBytes,
                MaxPartCount = 64,
                MaxPartUncompressedBytes = MaximumCharactersInPart,
                MaxTotalUncompressedBytes = MaximumPackageBytes,
                MaxCompressionRatio = 100D,
                Macros = OfficePackageContentPolicy.Reject,
                EmbeddedPayloads = OfficePackageContentPolicy.Reject,
                ActiveX = OfficePackageContentPolicy.Reject,
                ExternalRelationships = OfficePackageContentPolicy.Reject
            };

        private static byte[] ReadBounded(Stream stream) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));

            using var buffer = new MemoryStream();
            var chunk = new byte[81920];
            int totalBytes = 0;
            while (true) {
                int remaining = MaximumPackageBytes + 1 - totalBytes;
                int read = stream.Read(chunk, 0, Math.Min(chunk.Length, remaining));
                if (read == 0) {
                    return buffer.ToArray();
                }

                totalBytes = checked(totalBytes + read);
                if (totalBytes > MaximumPackageBytes) {
                    throw new InvalidDataException(
                        $"Embedded chart workbook exceeds the configured maximum of {MaximumPackageBytes} bytes.");
                }

                buffer.Write(chunk, 0, read);
            }
        }
    }
}
