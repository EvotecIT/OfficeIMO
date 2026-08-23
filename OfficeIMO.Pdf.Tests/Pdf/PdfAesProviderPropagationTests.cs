using OfficeIMO.Pdf;
using OfficeIMO.Security;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfAesProviderPropagationTests {
    [Fact]
    public void GeneratedEncryptedDocumentPreservesProviderForReadbackAndComplianceArtifact() {
        var provider = new CountingAesProvider();
        var encryption = new PdfStandardEncryptionOptions("open") {
            OwnerPassword = "owner",
            Algorithm = PdfStandardEncryptionAlgorithm.Aes256,
            AesCryptographyProvider = provider
        };
        PdfDocument document = PdfDocument.Create(new PdfOptions().SetEncryption(encryption))
            .Paragraph(paragraph => paragraph.Text("Managed provider readback."));

        _ = document.ToBytes();
        Assert.Same(provider, document.ReadOptions.AesCryptographyProvider);

        int readbackCount = provider.DecryptOperations;
        Assert.Contains("Managed provider readback.", document.Read.Text(), StringComparison.Ordinal);
        Assert.True(provider.DecryptOperations > readbackCount);

        PdfComplianceArtifact artifact = document.CreateComplianceArtifact(PdfComplianceProfile.PdfA3B);
        int complianceCount = provider.DecryptOperations;
        _ = artifact.AssessProof();
        Assert.True(provider.DecryptOperations > complianceCount);
    }

    [Fact]
    public void ReencryptUsesTheOutputProviderForValidationAndReturnedDocument() {
        var sourceProvider = new CountingAesProvider();
        var outputProvider = new CountingAesProvider();
        var sourceEncryption = new PdfStandardEncryptionOptions("source-open") {
            OwnerPassword = "source-owner",
            Algorithm = PdfStandardEncryptionAlgorithm.Aes256,
            AesCryptographyProvider = sourceProvider
        };
        byte[] source = PdfDocument.Create(new PdfOptions().SetEncryption(sourceEncryption))
            .Paragraph(paragraph => paragraph.Text("Provider replacement proof."))
            .ToBytes();
        var outputEncryption = new PdfStandardEncryptionOptions("output-open") {
            OwnerPassword = "output-owner",
            Algorithm = PdfStandardEncryptionAlgorithm.Aes256,
            AesCryptographyProvider = outputProvider
        };

        PdfSecurityMutationResult result = PdfSecurityEditor.Reencrypt(
            source,
            "source-owner",
            outputEncryption,
            new PdfReadOptions {
                Password = "source-owner",
                AesCryptographyProvider = sourceProvider
            });

        Assert.Same(outputProvider, result.OutputReadOptions?.AesCryptographyProvider);
        Assert.True(outputProvider.DecryptOperations > 0);
        int readbackCount = outputProvider.DecryptOperations;
        Assert.Contains("Provider replacement proof.", result.ToDocument().Read.Text(), StringComparison.Ordinal);
        Assert.True(outputProvider.DecryptOperations > readbackCount);
    }

    [Fact]
    public void EmptyPasswordRewriteOptionsPreserveTheSuppliedProvider() {
        var provider = new CountingAesProvider();
        var options = new PdfReadOptions {
            Password = "original",
            AesCryptographyProvider = provider
        };

        PdfReadOptions emptyPasswordOptions = PdfReadOptions.WithPassword(options, string.Empty);

        Assert.Equal(string.Empty, emptyPasswordOptions.Password);
        Assert.Same(provider, emptyPasswordOptions.AesCryptographyProvider);
    }

    private sealed class CountingAesProvider : IOfficeAesCryptographyProvider {
        public int EncryptOperations { get; private set; }
        public int DecryptOperations { get; private set; }
        public string Name => "Counting managed AES";

        public byte[] EncryptCbc(
            byte[] key,
            byte[] initializationVector,
            byte[] plaintext,
            OfficeAesPadding padding) {
            EncryptOperations++;
            return OfficeManagedAesCryptographyProvider.Default.EncryptCbc(
                key,
                initializationVector,
                plaintext,
                padding);
        }

        public byte[] DecryptCbc(
            byte[] key,
            byte[] initializationVector,
            byte[] ciphertext,
            OfficeAesPadding padding) {
            DecryptOperations++;
            return OfficeManagedAesCryptographyProvider.Default.DecryptCbc(
                key,
                initializationVector,
                ciphertext,
                padding);
        }
    }
}
