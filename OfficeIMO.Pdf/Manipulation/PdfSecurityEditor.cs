using OfficeIMO.Drawing.Internal;
namespace OfficeIMO.Pdf;

/// <summary>Adds, removes, or replaces Standard password security on supported existing PDFs.</summary>
internal static class PdfSecurityEditor {
    /// <summary>Encrypts an unencrypted PDF using modern Standard security by default.</summary>
    public static PdfSecurityMutationResult Encrypt(byte[] pdf, PdfStandardEncryptionOptions encryption) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(encryption, nameof(encryption));
        PdfDocumentSecurityInfo sourceSecurity = PdfSyntax.ReadDocumentSecurityInfo(pdf);
        if (sourceSecurity.HasEncryption) {
            throw new InvalidOperationException("The source PDF is already encrypted. Use Reencrypt with the owner password to replace its security settings.");
        }

        return Rewrite(pdf, sourceReadOptions: null, encryption, PdfSecurityMutationKind.Encrypt);
    }

    /// <summary>Removes Standard password security after authenticating the supplied owner password.</summary>
    public static PdfSecurityMutationResult Decrypt(byte[] pdf, string ownerPassword) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(ownerPassword, nameof(ownerPassword));
        return Rewrite(
            pdf,
            new PdfReadOptions { Password = ownerPassword },
            outputEncryption: null,
            PdfSecurityMutationKind.Decrypt);
    }

    /// <summary>Replaces Standard password security after authenticating the supplied current owner password.</summary>
    public static PdfSecurityMutationResult Reencrypt(
        byte[] pdf,
        string currentOwnerPassword,
        PdfStandardEncryptionOptions newEncryption) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(currentOwnerPassword, nameof(currentOwnerPassword));
        Guard.NotNull(newEncryption, nameof(newEncryption));
        return Rewrite(
            pdf,
            new PdfReadOptions { Password = currentOwnerPassword },
            newEncryption,
            PdfSecurityMutationKind.Reencrypt);
    }

    /// <summary>Encrypts an unencrypted PDF file and writes the proven rewrite to a new path.</summary>
    public static PdfSecurityMutationResult Encrypt(
        string inputPath,
        string outputPath,
        PdfStandardEncryptionOptions encryption) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNullOrWhiteSpace(outputPath, nameof(outputPath));
        PdfSecurityMutationResult result = Encrypt(File.ReadAllBytes(inputPath), encryption);
        OfficeFileCommit.WriteAllBytes(outputPath, result.Pdf);
        return result;
    }

    /// <summary>Decrypts a Standard-security PDF file and writes the proven rewrite to a new path.</summary>
    public static PdfSecurityMutationResult Decrypt(string inputPath, string outputPath, string ownerPassword) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNullOrWhiteSpace(outputPath, nameof(outputPath));
        PdfSecurityMutationResult result = Decrypt(File.ReadAllBytes(inputPath), ownerPassword);
        OfficeFileCommit.WriteAllBytes(outputPath, result.Pdf);
        return result;
    }

    private static PdfSecurityMutationResult Rewrite(
        byte[] sourcePdf,
        PdfReadOptions? sourceReadOptions,
        PdfStandardEncryptionOptions? outputEncryption,
        PdfSecurityMutationKind kind) {
        PdfMutationPlan plan = PdfMutationPlanner.RequireFullRewrite(
            sourcePdf,
            PdfMutationOperation.ChangeEncryption,
            sourceReadOptions);

        PdfDocumentSecurityInfo sourceSecurity = plan.Preflight.Probe.Security;
        ValidateSourceSecurity(kind, sourceSecurity);
        byte[] rewrittenPdf = PdfDocumentObjectGraphRewriter.Rewrite(sourcePdf, sourceReadOptions, outputEncryption);
        PdfReadOptions? outputReadOptions = outputEncryption is null
            ? null
            : new PdfReadOptions { Password = outputEncryption.OwnerPassword ?? outputEncryption.UserPassword };

        var preservationOptions = new PdfRewritePreservationOptions {
            OriginalReadOptions = sourceReadOptions,
            RewrittenReadOptions = outputReadOptions,
            PreserveDocumentVersionState = false,
            PreserveRevisionStructure = false,
            PreserveSecurityState = false
        };
        PdfRewritePreservationReport preservation = PdfRewritePreservation.Assess(sourcePdf, rewrittenPdf, preservationOptions);
        preservation.ThrowIfFailed();

        PdfDocumentSecurityInfo outputSecurity = PdfSyntax.ReadDocumentSecurityInfo(rewrittenPdf, outputReadOptions);
        ValidateOutputSecurity(kind, outputSecurity, outputEncryption);
        return new PdfSecurityMutationResult(
            kind,
            rewrittenPdf,
            plan,
            preservation,
            sourceSecurity,
            outputSecurity,
            outputReadOptions);
    }

    private static void ValidateSourceSecurity(PdfSecurityMutationKind kind, PdfDocumentSecurityInfo sourceSecurity) {
        if (kind == PdfSecurityMutationKind.Encrypt) {
            if (sourceSecurity.HasEncryption) {
                throw new InvalidOperationException("Encrypt requires an unencrypted source PDF.");
            }

            return;
        }

        if (!sourceSecurity.HasEncryption) {
            throw new InvalidOperationException(kind + " requires a Standard password-encrypted source PDF.");
        }

        if (!sourceSecurity.HasOwnerAuthorization) {
            throw new InvalidOperationException(kind + " requires the current owner password; a user password cannot authorize an encryption change.");
        }
    }

    private static void ValidateOutputSecurity(
        PdfSecurityMutationKind kind,
        PdfDocumentSecurityInfo outputSecurity,
        PdfStandardEncryptionOptions? outputEncryption) {
        bool shouldBeEncrypted = outputEncryption is not null;
        if (outputSecurity.HasEncryption != shouldBeEncrypted) {
            throw new InvalidOperationException("The rewritten PDF security state did not match the requested " + kind + " operation.");
        }

        if (shouldBeEncrypted &&
            outputSecurity.PasswordAuthenticationRole == PdfPasswordAuthenticationRole.None) {
            throw new InvalidOperationException("The rewritten encrypted PDF could not be opened with its configured user password.");
        }
    }
}
