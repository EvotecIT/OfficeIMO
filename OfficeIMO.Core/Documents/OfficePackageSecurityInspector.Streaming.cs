using OfficeIMO.Core.Internal;

namespace OfficeIMO {
    public static partial class OfficePackageSecurityInspector {
        // The loader and parser use the same fixed window. Capture-only callers
        // receive findings without turning them into admission failures.
        internal static OfficePackageSecurityReport InspectSeekableForLoad(Stream source,
            OfficePackageSecurityOptions? policy, CancellationToken cancellationToken) {
            var report = InspectSeekableSource(source, policy ?? OfficePackageSecurityOptions.SecureDefaults,
                cancellationToken, reportOversizedPackage: policy == null);
            if (policy != null) ThrowFirstError(report);
            return report;
        }

        private static OfficePackageSecurityReport InspectSeekableSource(
            Stream source,
            OfficePackageSecurityOptions options,
            CancellationToken cancellationToken = default,
            bool reportOversizedPackage = false) {
            cancellationToken.ThrowIfCancellationRequested();
            ValidateOptions(options);
            if (!source.CanRead || !source.CanSeek) {
                throw new ArgumentException(
                    "Streaming package inspection requires a readable seekable stream.",
                    nameof(source));
            }

            long originalPosition = source.Position;
            try {
                long sourceLength = source.Length;
                if (originalPosition < 0 || originalPosition > sourceLength) {
                    throw new ArgumentException(
                        "The source position must be within the readable stream length.",
                        nameof(source));
                }
                long packageBytes = sourceLength;
                if (!reportOversizedPackage) ValidateSourceSize(packageBytes, options);
                var findings = new System.Collections.Generic.List<OfficePackageSecurityFinding>();

                using var package = new SeekableReadWindowStream(source, 0, packageBytes, cancellationToken: cancellationToken);
                var signature = new byte[8];
                int signatureBytes = ReadPrefix(package, signature);
                package.Position = 0;
                bool isZip = HasZipSignature(signature, signatureBytes);
                bool isCompound = OfficeIMO.Core.Internal.OfficeCompoundDocumentDetector
                    .HasCompoundSignature(signature);
                if (packageBytes > options.MaxPackageBytes) {
                    AddLimitFinding(findings, OfficePackageSecurityRule.PackageSize, packageBytes,
                        options.MaxPackageBytes, "Package size");
                    OfficePackageContainerKind kind = isZip ? OfficePackageContainerKind.OpenXml
                        : isCompound ? OfficePackageContainerKind.CompoundBinary : OfficePackageContainerKind.Unknown;
                    return new OfficePackageSecurityReport(packageBytes, kind,
                        0, 0, 0, 0, 0, 0, 0, 0, 0, findings.ToArray());
                }
                if (isZip) return InspectZip(package, packageBytes, options, findings);
                if (isCompound) return InspectCompound(package, packageBytes, options, findings);
                return new OfficePackageSecurityReport(packageBytes, OfficePackageContainerKind.Unknown,
                    0, 0, 0, 0, 0, 0, 0, 0, 0, findings.ToArray());
            } finally {
                source.Position = originalPosition;
            }
        }

        private static int ReadPrefix(Stream source, byte[] buffer) {
            int total = 0;
            while (total < buffer.Length) {
                int read = source.Read(buffer, total, buffer.Length - total);
                if (read == 0) break;
                total += read;
            }
            return total;
        }

        private static bool HasZipSignature(byte[] bytes, int length) => length >= 4
            && bytes[0] == 0x50 && bytes[1] == 0x4b
            && ((bytes[2] == 0x03 && bytes[3] == 0x04)
                || (bytes[2] == 0x05 && bytes[3] == 0x06)
                || (bytes[2] == 0x07 && bytes[3] == 0x08));

    }
}
