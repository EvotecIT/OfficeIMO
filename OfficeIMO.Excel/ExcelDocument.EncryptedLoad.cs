using System.Threading;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Excel;

public partial class ExcelDocument {
    private static ExcelDocument LoadEncryptedInput(byte[] encryptedBytes, string password,
        ExcelLoadOptions options, string? sourcePath, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (options.PackageSecurity != null) {
            OfficePackageSecurityInspector.Validate(encryptedBytes, options.PackageSecurity);
        }

        ExcelDocument document;
        if (ExcelDocumentLoadRouting.IsEncryptedLegacyXls(encryptedBytes, sourcePath)) {
            document = LoadEncryptedLegacyXls(encryptedBytes, password,
                options.AccessMode == DocumentAccessMode.ReadOnly, options, cancellationToken);
        } else {
            byte[] packageBytes = OfficeEncryption.DecryptPackage(encryptedBytes, password, cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            document = LoadFromByteArray(packageBytes, options, filePath: null, captureSourcePackageSecurity: false);
        }

        document = WithSourcePackageInspection(document, encryptedBytes, options, "encrypted-source-container");
        if (cancellationToken.IsCancellationRequested) {
            // The loaded package is detached and uses explicit persistence, so cancellation cannot save it.
            document.Dispose();
            cancellationToken.ThrowIfCancellationRequested();
        }
        return document;
    }
}
