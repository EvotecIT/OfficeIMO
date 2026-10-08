using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Benchmarks {
    /// <summary>Opt-in search, review and verified-output workload. Input preparation and saved-output readback are outside measurement.</summary>
    public sealed class PdfRedactionRuntimeWorkload {
        private readonly byte[] _input;
        private readonly string _pattern;
        private PdfRedactionPlan? _plan;
        private PdfRedactionApplyResult? _result;

        public PdfRedactionRuntimeWorkload(int pages, int rotation = 0) {
            if (pages < 1 || pages > 500) throw new ArgumentOutOfRangeException(nameof(pages));
            PdfDocument document = PdfDocument.Create(compose => {
                for (int page = 1; page <= pages; page++) {
                    int number = page;
                    compose.Page(layout => layout.Size(600, 800).Content(content => content.Item(item =>
                        item.Paragraph(text => text.Text($"Before private account {number:000} after page {number:000}")))));
                }
            });
            if (rotation != 0) document = document.Pages.Rotate(rotation, Enumerable.Range(1, pages).ToArray());
            _input = document.ToBytes();
            _pattern = @"private account [0-9]{3}";
            ExpectedPages = pages;
            SourceSha256 = Convert.ToHexString(SHA256.HashData(_input));
        }

        public PdfRedactionRuntimeWorkload(string inputPath, string pattern) {
            _input = File.ReadAllBytes(inputPath);
            _pattern = pattern;
            ExpectedPages = PdfReadDocument.Open(_input).Pages.Count;
            SourceSha256 = Convert.ToHexString(SHA256.HashData(_input));
        }

        public int ExpectedPages { get; }
        public int VerifiedPages { get; private set; }
        public int SelectedAreas => _plan?.Areas.Count ?? 0;
        public long InputBytes => _input.LongLength;
        public long OutputBytes => _result?.Pdf.LongLength ?? 0;
        public string SourceSha256 { get; }

        /// <summary>Measures the same owner calls as Studio: search, relabel/review, apply and verify including managed rendering.</summary>
        public void Execute() {
            PdfDocument source = PdfDocument.Load(_input);
            PdfRedactionPlan search = source.Redactions.Search(new PdfRedactionSearchOptions {
                TextSelection = PdfRedactionTextSelection.MatchedGlyphs,
                ContentScope = PdfRedactionContentScope.TextOnly,
                MatchCase = true
            }.AddRegex(_pattern));
            if (!search.IsReviewable || search.Areas.Count == 0) {
                throw new InvalidOperationException("Search is blocked or empty: " + string.Join(" ", search.Findings.Select(finding => finding.Code)));
            }
            _plan = source.Redactions.Plan(search.Areas.Select(area => area.WithLabel("Reviewed")));
            _result = source.Redactions.ApplyWithEvidence(_plan, verificationOptions: new PdfRedactionVerificationOptions {
                CheckManagedRendering = true,
                RequireCompleteStreamInspection = true,
                FailOnUndecodablePdfStreams = true
            }).ThrowIfUnverified();
            if (!_input.AsSpan().SequenceEqual(source.ToBytes())) throw new InvalidDataException("Source was modified.");
        }

        /// <summary>Checks the saved artifact, page coverage, absence and neighboring text after the measured operation.</summary>
        public void Validate() {
            if (_result is null || _plan is null || !_result.Evidence.IsVerified) throw new InvalidDataException("No verified result.");
            PdfReadDocument read = PdfReadDocument.Open(_result.Pdf);
            if (read.Pages.Count != ExpectedPages) throw new InvalidDataException("Page count changed.");
            var expression = new System.Text.RegularExpressions.Regex(_pattern,
                System.Text.RegularExpressions.RegexOptions.CultureInvariant, TimeSpan.FromSeconds(2));
            foreach (PdfReadPage page in read.Pages) {
                string text = page.ExtractText();
                if (expression.IsMatch(text)) throw new InvalidDataException("Selected text survives.");
                if (_pattern == @"private account [0-9]{3}" &&
                    (!text.Contains("Before", StringComparison.Ordinal) || !text.Contains("after page", StringComparison.Ordinal))) {
                    throw new InvalidDataException("Neighboring text was removed.");
                }
            }
            VerifiedPages = read.Pages.Count;
            if (!string.Equals(SourceSha256, Convert.ToHexString(SHA256.HashData(_input)), StringComparison.Ordinal))
                throw new InvalidDataException("Input fingerprint changed.");
        }

        /// <summary>Retains a compact source/result pair and exact reviewed rectangles for independent viewer checks.</summary>
        public void Export(string directory) {
            Validate();
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(Path.Combine(directory, "source.pdf"), _input);
            File.WriteAllBytes(Path.Combine(directory, "redacted.pdf"), _result!.Pdf);
            File.WriteAllText(Path.Combine(directory, "evidence.json"), JsonSerializer.Serialize(new {
                SourceSha256,
                OutputSha256 = Convert.ToHexString(SHA256.HashData(_result.Pdf)),
                VerifiedPages,
                Areas = _plan!.Areas.Select(area => new { area.PageNumber, area.X, area.Y, area.Width, area.Height })
            }, new JsonSerializerOptions { WriteIndented = true }));
        }

        public void ReleaseResults() {
            _plan = null;
            _result = null;
        }
    }
}
