using OfficeIMO.Pdf;
using System.Text.RegularExpressions;
using Xunit;

namespace OfficeIMO.Tests.Pdf {
    public sealed class PdfRedactionImportedCorpusTests {
        [Theory]
        [InlineData("importedcarlito-rotation-0.pdf")]
        [InlineData("embedded-cff-rotation-0.pdf")]
        [InlineData("imported-truetype-and-ink.pdf")]
        public void ImportedFontsSupportVerifiedPreciseRemovalAndPreservePublicContent(string fixture) {
            byte[] bytes = ReadFixture(fixture);
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionPlan plan = source.Redactions.Search(new PdfRedactionSearchOptions {
                TextSelection = PdfRedactionTextSelection.MatchedGlyphs,
                ContentScope = PdfRedactionContentScope.TextOnly
            }.AddLiteral("private account 123"));
            Assert.True(plan.IsReviewable, string.Join(" ", plan.Findings.Select(finding => finding.Message)));
            Assert.Equal(source.Read().Pages.Count, plan.Areas.Count);

            PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(plan);

            Assert.True(result.Evidence.IsVerified, result.Evidence.Summary);
            PdfReadDocument read = PdfReadDocument.Open(result.Pdf);
            foreach (PdfReadPage page in read.Pages) {
                Assert.DoesNotMatch(@"private\s+account\s+123", page.ExtractText());
                Assert.Contains("Before", page.ExtractText(), StringComparison.Ordinal);
                Assert.Contains("after page", page.ExtractText(), StringComparison.Ordinal);
            }
            AssertRetainedNavigationAndInk(PdfReadDocument.Open(bytes), read);
            Assert.Equal(bytes, source.ToBytes());
        }

        [Fact]
        public void ReviewedAreaRemovesOutlinedLettersAndPreservesInkAndBookmark() {
            byte[] bytes = ReadFixture("outline-text-and-ink.pdf");
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionPlan plan = source.Redactions.Plan(new[] {
                new PdfRedactionArea(1, 70D, 598D, 150.388671875D, 26D, "Reviewed outline")
            });
            Assert.Contains(plan.Matches, match => match.Kind == PdfRedactionMatchKind.VectorPath);

            PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(plan);

            Assert.True(result.Evidence.IsVerified, result.Evidence.Summary);
            PdfReadDocument read = PdfReadDocument.Open(result.Pdf);
            Assert.Contains("Before public summary after page", read.Pages[0].ExtractText(), StringComparison.Ordinal);
            AssertRetainedNavigationAndInk(PdfReadDocument.Open(bytes), read);
            Assert.Equal(bytes, source.ToBytes());
        }

        private static byte[] ReadFixture(string name) => File.ReadAllBytes(Path.Combine(
            AppContext.BaseDirectory, "Pdf", "Fixtures", "Interoperability", "Redaction", name));

        private static void AssertRetainedNavigationAndInk(PdfReadDocument before, PdfReadDocument after) {
            Assert.Equal(before.Outlines.Select(item => item.Title), after.Outlines.Select(item => item.Title));
            for (int page = 0; page < before.Pages.Count; page++) {
                PdfAnnotation[] sourceInk = before.Pages[page].GetAnnotations().Where(item => item.Subtype == "Ink").ToArray();
                PdfAnnotation[] outputInk = after.Pages[page].GetAnnotations().Where(item => item.Subtype == "Ink").ToArray();
                Assert.Equal(sourceInk.Length, outputInk.Length);
                for (int index = 0; index < sourceInk.Length; index++) {
                    Assert.Equal(sourceInk[index].Contents, outputInk[index].Contents);
                    Assert.Equal(sourceInk[index].InkList.Count, outputInk[index].InkList.Count);
                    for (int stroke = 0; stroke < sourceInk[index].InkList.Count; stroke++) {
                        Assert.Equal(sourceInk[index].InkList[stroke], outputInk[index].InkList[stroke]);
                    }
                }
            }
        }
    }
}