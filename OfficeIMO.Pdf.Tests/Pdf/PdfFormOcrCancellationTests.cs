using System.Collections;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Ocr;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using OfficeIMO.Security;
using Xunit;

namespace OfficeIMO.Tests.Pdf {
    public sealed class PdfFormOcrCancellationTests {
        [Theory]
        [InlineData(false, false, false, true)]
        [InlineData(true, false, false, true)]
        [InlineData(true, true, true, false)]
        [InlineData(false, true, true, false)]
        [InlineData(true, false, false, false)]
        [InlineData(true, false, true, false)]
        public async Task ReviewedFormValuesObserveCancellationWithoutChangingAuthorizedMutationModes(
            bool acceptValue, bool cancelDuringDecryption, bool encrypted, bool cancelAfterAcceptance) {
            using var cancellation = new CancellationTokenSource();
            var provider = new CancellationTrackingAesProvider(cancellation.Token);
            var (source, original, review) = await CreateReviewAsync(provider, encrypted);
            var values = new Dictionary<PdfFormOcrProposal, PdfFormFieldValue>();
            if (acceptValue) values.Add(review.Proposals.Single(), "Reviewed");
            IReadOnlyDictionary<PdfFormOcrProposal, PdfFormFieldValue> accepted = cancelAfterAcceptance
                ? new CancellingAcceptance(values, cancellation) : values;
            if (cancelDuringDecryption) provider.CancelOnNextDecryption = cancellation;

            if (cancelAfterAcceptance || cancelDuringDecryption) {
                var exception = Assert.ThrowsAny<OperationCanceledException>(() => review.Apply(source, accepted, cancellation.Token));
                Assert.Equal(cancellation.Token, exception.CancellationToken);
            } else {
                PdfDocument result = review.Apply(source, accepted, cancellation.Token);
                Assert.Equal("Reviewed", result.Inspect().FormFieldsByName["Name"].Value);
                byte[] output = result.ToBytes();
                Assert.Equal(encrypted, output.Length >= original.Length && output.AsSpan(0, original.Length).SequenceEqual(original));
            }
            Assert.Equal(0, provider.EncryptionsAfterCancellation);
            Assert.Equal(0, provider.DecryptionsAfterCancellation);
            Assert.Equal(original, source.ToBytes());
        }

        [Fact]
        public async Task AppendEncryptionCancellationStopsLaterEncryptionAndOutputReadback() {
            using var cancellation = new CancellationTokenSource();
            var provider = new CancellationTrackingAesProvider(cancellation.Token);
            var (source, original, review) = await CreateReviewAsync(provider, encrypted: true);
            var accepted = new Dictionary<PdfFormOcrProposal, PdfFormFieldValue> {
                [review.Proposals.Single()] = "Reviewed"
            };
            provider.CancelOnNextEncryption = cancellation;

            var exception = Assert.ThrowsAny<OperationCanceledException>(() => review.Apply(source, accepted, cancellation.Token));

            Assert.Equal(cancellation.Token, exception.CancellationToken);
            Assert.Equal(0, provider.EncryptionsAfterCancellation);
            Assert.Equal(0, provider.DecryptionsAfterCancellation);
            Assert.Equal(original, source.ToBytes());
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public async Task AsyncProtectedLoadObservesCancellationDuringArtifactDecryption(bool fromPath) {
            using var cancellation = new CancellationTokenSource();
            var provider = new CancellationTrackingAesProvider(cancellation.Token);
            byte[] original = CreatePdfBytes(encrypted: true);
            var readOptions = new PdfLoadOptions { Password = "owner", AesCryptographyProvider = provider };
            using var stream = new MemoryStream(original);
            stream.Position = 7;
            string? fixturePath = null;
            try {
                if (fromPath) {
                    fixturePath = Path.Combine(Path.GetTempPath(), "officeimo-pdf-async-cancellation-" + Guid.NewGuid().ToString("N") + ".pdf");
                    File.WriteAllBytes(fixturePath, original);
                }
                provider.CancelOnNextDecryption = cancellation;

                var exception = await Assert.ThrowsAnyAsync<OperationCanceledException>(() => fromPath
                    ? PdfDocument.LoadAsync(fixturePath!, readOptions, cancellation.Token)
                    : PdfDocument.LoadAsync(stream, readOptions, cancellation.Token));

                Assert.Equal(cancellation.Token, exception.CancellationToken);
                Assert.Equal(0, provider.EncryptionsAfterCancellation);
                Assert.Equal(0, provider.DecryptionsAfterCancellation);
                Assert.Equal(7, stream.Position);
                Assert.Equal(original, fromPath ? File.ReadAllBytes(fixturePath!) : stream.ToArray());
            } finally {
                if (fixturePath is not null) File.Delete(fixturePath);
            }
        }

        [Fact]
        public void FullRewriteCancellationPrecedesNewAppearanceFontWork() {
            byte[] original = CreatePdfBytes(encrypted: false);
            byte[] source = original.ToArray();
            var options = new PdfFormFillerOptions().UseAppearanceFont("Unavailable appearance fixture", [0, 1, 2, 3]);
            var values = new Dictionary<string, PdfFormFieldValue> { ["Name"] = "Reviewed" };
            Assert.Throws<InvalidOperationException>(() => PdfFormFiller.FillFields(source, values, options, readOptions: null));
            using var cancellation = new CancellationTokenSource();
            var accepted = new CancellingFieldValues(values, cancellation);

            var exception = Assert.ThrowsAny<OperationCanceledException>(() =>
                PdfFormFiller.FillFields(source, accepted, options, readOptions: null, cancellation.Token));

            Assert.Equal(cancellation.Token, exception.CancellationToken);
            Assert.Equal(original, source);
        }

        [Theory]
        [InlineData("owner", 2)]
        [InlineData("open", 3)]
        public async Task ModernAuthenticationCancellationStopsFileKeyDecryption(string password, int completedHashes) {
            byte[] original = PdfDocument.Create(new PdfOptions().SetEncryption(new PdfStandardEncryptionOptions("open") {
                OwnerPassword = "owner", Algorithm = PdfStandardEncryptionAlgorithm.Aes256
            })).TextField("Name", width: 180, height: 24).ToBytes();
            using var cancellation = new CancellationTokenSource();
            var provider = new CancellationTrackingAesProvider(cancellation.Token) {
                Revision6HashToCancel = completedHashes, CancelOnCompletedRevision6Hash = cancellation
            };
            using var stream = new MemoryStream(original);
            stream.Position = 7;

            var exception = await Assert.ThrowsAnyAsync<OperationCanceledException>(() => PdfDocument.LoadAsync(stream,
                new PdfLoadOptions { Password = password, AesCryptographyProvider = provider }, cancellation.Token));

            Assert.Equal(completedHashes, provider.CompletedRevision6Hashes);
            Assert.Equal(cancellation.Token, exception.CancellationToken);
            Assert.Equal(0, provider.EncryptionsAfterCancellation);
            Assert.Equal(0, provider.DecryptionsAfterCancellation);
            Assert.Equal(7, stream.Position);
            Assert.Equal(original, stream.ToArray());
        }

        [Fact]
        public void Aes256GenerationCancellationStopsPermissionsEncryption() {
            using var cancellation = new CancellationTokenSource();
            var provider = new CancellationTrackingAesProvider(cancellation.Token) { CancelOnOwnerFileKeyEncryption = cancellation };
            var source = PdfDocument.Create(new PdfOptions().SetEncryption(new PdfStandardEncryptionOptions("open") {
                OwnerPassword = "owner", Algorithm = PdfStandardEncryptionAlgorithm.Aes256, AesCryptographyProvider = provider
            })).TextField("Name", width: 180, height: 24);

            var exception = Assert.ThrowsAny<OperationCanceledException>(() => source.ToBytes(cancellation.Token));

            Assert.Equal(2, provider.FileKeyEncryptions);
            Assert.Equal(cancellation.Token, exception.CancellationToken);
            Assert.Equal(0, provider.EncryptionsAfterCancellation);
            Assert.Equal(0, provider.DecryptionsAfterCancellation);
        }

        private static async Task<(PdfDocument Source, byte[] Bytes, PdfFormOcrReview Review)> CreateReviewAsync(
            CancellationTrackingAesProvider provider, bool encrypted) {
            byte[] original = CreatePdfBytes(encrypted);
            var source = PdfDocument.Load(original, new PdfLoadOptions {
                Password = encrypted ? "owner" : null, AesCryptographyProvider = provider
            });
            // The canonical planner prefers an available append-only form update for authenticated encrypted input.
            Assert.Equal(encrypted ? PdfMutationExecutionMode.AppendOnly : PdfMutationExecutionMode.FullRewrite,
                source.PlanMutation(PdfMutationOperation.FillFormFields, ["Name"]).ExecutionMode);
            var widget = source.Inspect().FormFieldsByName["Name"].Widgets.Single();
            var bounds = source.GetPageLayouts()[0].MapUserSpaceRectangleToVisual(widget.X1, widget.Y1, widget.X2, widget.Y2);
            var engine = new DelegateOcrEngine("cancellation-fixture", (_, _) => Task.FromResult(new OcrResult {
                Spans = [new OcrTextSpan {
                    Text = "Reviewed", Confidence = .99, Level = OcrTextSpanLevel.Word, CoordinateUnit = OcrCoordinateUnit.Points,
                    Region = new() { X = bounds.Left + 1, Y = bounds.Top + 1, Width = bounds.Width - 2, Height = bounds.Height - 2 }
                }]
            }));
            var review = await source.PrepareFormOcrAsync(engine, new() { Dpi = 72 });
            return (source, original, review);
        }

        private static byte[] CreatePdfBytes(bool encrypted) {
            var options = new PdfOptions();
            if (encrypted) options.SetEncryption(new PdfStandardEncryptionOptions("open") {
                OwnerPassword = "owner", Algorithm = PdfStandardEncryptionAlgorithm.Aes128
            });
            return PdfDocument.Create(options).TextField("Name", width: 180, height: 24).ToBytes();
        }

        // A caller-owned selection cancels at its completion boundary, without a timer or a production hook.
        private sealed class CancellingAcceptance : IReadOnlyDictionary<PdfFormOcrProposal, PdfFormFieldValue> {
            private readonly IReadOnlyDictionary<PdfFormOcrProposal, PdfFormFieldValue> _values;
            private readonly CancellationTokenSource _cancellation;

            public CancellingAcceptance(IReadOnlyDictionary<PdfFormOcrProposal, PdfFormFieldValue> values,
                CancellationTokenSource cancellation) {
                _values = values;
                _cancellation = cancellation;
            }

            public int Count => _values.Count;
            public IEnumerable<PdfFormOcrProposal> Keys => _values.Keys;
            public IEnumerable<PdfFormFieldValue> Values => _values.Values;
            public PdfFormFieldValue this[PdfFormOcrProposal key] => _values[key];
            public bool ContainsKey(PdfFormOcrProposal key) => _values.ContainsKey(key);
            public bool TryGetValue(PdfFormOcrProposal key, out PdfFormFieldValue value) => _values.TryGetValue(key, out value!);

            public IEnumerator<KeyValuePair<PdfFormOcrProposal, PdfFormFieldValue>> GetEnumerator() {
                foreach (var value in _values) yield return value;
                _cancellation.Cancel();
            }

            IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
        }

        // A supported caller dictionary can cancel when the canonical field owner resolves its value.
        private sealed class CancellingFieldValues : IReadOnlyDictionary<string, PdfFormFieldValue> {
            private readonly IReadOnlyDictionary<string, PdfFormFieldValue> _values;
            private readonly CancellationTokenSource _cancellation;
            public CancellingFieldValues(IReadOnlyDictionary<string, PdfFormFieldValue> values, CancellationTokenSource cancellation) {
                _values = values;
                _cancellation = cancellation;
            }
            public int Count => _values.Count;
            public IEnumerable<string> Keys => _values.Keys;
            public IEnumerable<PdfFormFieldValue> Values => _values.Values;
            public PdfFormFieldValue this[string key] => _values[key];
            public bool ContainsKey(string key) => _values.ContainsKey(key);
            public bool TryGetValue(string key, out PdfFormFieldValue value) {
                bool found = _values.TryGetValue(key, out value!);
                _cancellation.Cancel();
                return found;
            }
            public IEnumerator<KeyValuePair<string, PdfFormFieldValue>> GetEnumerator() => _values.GetEnumerator();
            IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
        }

        // The supported AES boundary performs real decryption and exposes work begun after host cancellation.
        private sealed class CancellationTrackingAesProvider : IOfficeAesCryptographyProvider {
            private readonly CancellationToken _cancellation;

            public CancellationTrackingAesProvider(CancellationToken cancellation) => _cancellation = cancellation;
            public string Name => "Cancellation-tracking managed AES";
            public int EncryptionsAfterCancellation { get; private set; }
            public int DecryptionsAfterCancellation { get; private set; }
            public CancellationTokenSource? CancelOnNextEncryption { get; set; }
            public CancellationTokenSource? CancelOnNextDecryption { get; set; }
            public CancellationTokenSource? CancelOnCompletedRevision6Hash { get; set; }
            public int Revision6HashToCancel { get; set; }
            public int CompletedRevision6Hashes { get; private set; }
            public CancellationTokenSource? CancelOnOwnerFileKeyEncryption { get; set; }
            public int FileKeyEncryptions { get; private set; }
            private int _revision6Round;

            public byte[] EncryptCbc(byte[] key, byte[] initializationVector, byte[] plaintext, OfficeAesPadding padding) {
                if (_cancellation.IsCancellationRequested) EncryptionsAfterCancellation++;
                CancellationTokenSource? cancellation = CancelOnNextEncryption;
                CancelOnNextEncryption = null;
                cancellation?.Cancel();
                byte[] encrypted = OfficeManagedAesCryptographyProvider.Default.EncryptCbc(key, initializationVector, plaintext, padding);
                if (padding == OfficeAesPadding.None && key.Length == 16 && plaintext.Length >= 32 * 64) {
                    _revision6Round++;
                    if (_revision6Round >= 64 && encrypted[encrypted.Length - 1] <= _revision6Round - 32) {
                        _revision6Round = 0;
                        CompletedRevision6Hashes++;
                        if (CompletedRevision6Hashes == Revision6HashToCancel) CancelOnCompletedRevision6Hash?.Cancel();
                    }
                }
                if (padding == OfficeAesPadding.None && key.Length == 32 && plaintext.Length == 32) {
                    FileKeyEncryptions++;
                    if (FileKeyEncryptions == 2) CancelOnOwnerFileKeyEncryption?.Cancel();
                }
                return encrypted;
            }

            public byte[] DecryptCbc(byte[] key, byte[] initializationVector, byte[] ciphertext, OfficeAesPadding padding) {
                if (_cancellation.IsCancellationRequested) DecryptionsAfterCancellation++;
                CancellationTokenSource? cancellation = CancelOnNextDecryption;
                CancelOnNextDecryption = null;
                cancellation?.Cancel();
                return OfficeManagedAesCryptographyProvider.Default.DecryptCbc(key, initializationVector, ciphertext, padding);
            }
        }
    }
}
