using OfficeIMO.Pdf;
using OfficeIMO.Security;
using Xunit;
using System.Threading.Tasks;
using System.Threading;
using OfficeIMO.Drawing;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfAesProviderPropagationTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task WaitingForColdParseHonorsCancellationAndDisplayDeadline(bool useDeadline) {
        byte[] bytes = PdfDocument.Create(new PdfOptions().SetEncryption(new PdfStandardEncryptionOptions("open") {
            OwnerPassword = "owner", Algorithm = PdfStandardEncryptionAlgorithm.Aes128
        })).Paragraph(paragraph => paragraph.Text("Cold parse cancellation")).ToBytes();
        using var parsing = new ManualResetEventSlim();
        using var release = new ManualResetEventSlim();
        var provider = new CountingAesProvider();
        // Adopt an encrypted operation result with a cold canonical cache, as mutation paths do.
        var document = PdfDocument.Load(bytes, new PdfLoadOptions { Password = "owner", AesCryptographyProvider = provider })
            .WithBytes(bytes, bytes);
        provider.BeforeDecrypt = () => {
            parsing.Set();
            if (!release.Wait(TimeSpan.FromSeconds(30))) throw new TimeoutException("Test parser was not released.");
        };
        Task<PdfPageRenderResult> first = Task.Factory.StartNew(() => document.Render.DisplayPage(1,
            new PdfPageDisplayOptions { MaximumDimension = 80 }), CancellationToken.None, TaskCreationOptions.LongRunning, TaskScheduler.Default);
        Task<PdfPageRenderResult>? waiting = null;
        try {
            Assert.True(parsing.Wait(TimeSpan.FromSeconds(10)));
            using var cancellation = new CancellationTokenSource();
            waiting = Task.Factory.StartNew(() => {
                if (!useDeadline) cancellation.CancelAfter(TimeSpan.FromMilliseconds(200));
                return document.Render.DisplayPage(1, new PdfPageDisplayOptions {
                    MaximumDimension = 80, Timeout = useDeadline ? TimeSpan.FromMilliseconds(200) : null
                }, cancellation.Token);
            }, CancellationToken.None, TaskCreationOptions.LongRunning, TaskScheduler.Default);
            Assert.Same(waiting, await Task.WhenAny(waiting, Task.Delay(TimeSpan.FromSeconds(5))));
            Assert.False(first.IsCompleted);
            if (useDeadline) await Assert.ThrowsAsync<OfficeImageExportTimeoutException>(async () => await waiting);
            else await Assert.ThrowsAnyAsync<OperationCanceledException>(async () => await waiting);
        } finally {
            release.Set();
            await first;
            if (waiting is not null) { try { await waiting; } catch (Exception) { } }
        }
        Assert.Equal((await first).Bytes, document.Render.DisplayPage(1, new PdfPageDisplayOptions { MaximumDimension = 80 }).Bytes);
    }

    [Fact]
    public async Task OpenedDocumentNavigationReusesDecryptionAcrossDisplayAndEveryRenderSelection() {
        var provider = new CountingAesProvider();
        byte[] bytes = PdfDocument.Create(new PdfOptions().SetEncryption(new PdfStandardEncryptionOptions("open") {
            OwnerPassword = "owner", Algorithm = PdfStandardEncryptionAlgorithm.Aes128, AesCryptographyProvider = provider
        })).Paragraph(paragraph => paragraph.Text("First cached page")).PageBreak()
            .Paragraph(paragraph => paragraph.Text("Second cached page")).ToBytes();
        var options = new PdfLoadOptions { Password = "open", AesCryptographyProvider = provider };
        PdfDocument document = PdfDocument.Load(bytes, options);
        Assert.Equal(2, document.InspectGeometryForViewing().PageCount);
        int parsedDecryptions = provider.DecryptOperations;
        Assert.True(parsedDecryptions > 0);
        var display = new PdfPageDisplayOptions { MaximumDimension = 80 };
        byte[] first = document.Render.DisplayPage(1, display).Bytes!;
        byte[] second = document.Render.DisplayPage(2, display).Bytes!;
        var render = new PdfPageRenderOptions { ThumbnailMaxDimension = 80, ContinueOnError = false };
        Assert.Equal(second, document.Render.Pages("2", render)[0].Bytes);
        Assert.Equal(first, document.Render.Pages(PdfPageSelection.From(1), render)[0].Bytes);
        Assert.Equal(second, document.Reader.RenderPages(PdfPageSelector.Parse("last"), render)[0].Bytes);
        var concurrent = await Task.WhenAll(Enumerable.Range(0, 12).Select(index => Task.Run(() =>
            document.Render.DisplayPage(index % 2 + 1, display).Bytes)));
        for (int index = 0; index < concurrent.Length; index++) Assert.Equal(index % 2 == 0 ? first : second, concurrent[index]);
        Assert.Equal(parsedDecryptions, provider.DecryptOperations);

        using var cancellation = new System.Threading.CancellationTokenSource();
        cancellation.Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() => document.Render.DisplayPage(1, display, cancellation.Token));
        Assert.Equal(first, document.Render.DisplayPage(1, display).Bytes);
        Assert.Throws<PdfInvalidPasswordException>(() => document.Reader.RenderPages("1", render,
            new PdfLoadOptions { Password = "wrong", AesCryptographyProvider = provider }));
        Assert.Equal(first, document.Render.DisplayPage(1, display).Bytes);
    }

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
        Assert.Contains("Managed provider readback.", document.Reader.Text(), StringComparison.Ordinal);
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
            new PdfLoadOptions {
                Password = "source-owner",
                AesCryptographyProvider = sourceProvider
            });

        Assert.Same(outputProvider, result.OutputReadOptions?.AesCryptographyProvider);
        Assert.True(outputProvider.DecryptOperations > 0);
        int readbackCount = outputProvider.DecryptOperations;
        Assert.Contains("Provider replacement proof.", result.ToDocument().Reader.Text(), StringComparison.Ordinal);
        Assert.True(outputProvider.DecryptOperations > readbackCount);
    }

    [Fact]
    public void EmptyPasswordRewriteOptionsPreserveTheSuppliedProvider() {
        var provider = new CountingAesProvider();
        var options = new PdfLoadOptions {
            Password = "original",
            AesCryptographyProvider = provider
        };

        PdfLoadOptions emptyPasswordOptions = PdfLoadOptions.WithPassword(options, string.Empty);

        Assert.Equal(string.Empty, emptyPasswordOptions.Password);
        Assert.Same(provider, emptyPasswordOptions.AesCryptographyProvider);
    }

    private sealed class CountingAesProvider : IOfficeAesCryptographyProvider {
        internal Action? BeforeDecrypt { get; set; }
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
            BeforeDecrypt?.Invoke();
            DecryptOperations++;
            return OfficeManagedAesCryptographyProvider.Default.DecryptCbc(
                key,
                initializationVector,
                ciphertext,
                padding);
        }
    }
}
