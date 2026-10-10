using System.Threading.Tasks;
using OfficeIMO.Ocr;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using Xunit;

namespace OfficeIMO.Tests.Pdf
{
    public sealed class PdfFormOcrReviewBudgetTests {
        [Fact]
        public async Task AcceptedCorrectionsOnSeparatePagesUseIndependentCharacterBudgets() {
            PdfDocument source = PdfFormOcrReviewTests.Fixture("reportlab-scanned-form.pdf");
            byte[] original = source.ToBytes();
            PdfFormOcrReview review = await source.PrepareFormOcrAsync(PdfFormOcrReviewTests.Engine(source), new() {
                Dpi = 72, MaxOcrTextCharactersPerPage = 36
            });
            PdfFormOcrProposal name = review.Proposals.Single(proposal => proposal.Field.Name == "FullName");
            PdfFormOcrProposal reference = review.Proposals.Single(proposal => proposal.Field.Name == "Reference");

            PdfDocument result = review.Apply(source, new Dictionary<PdfFormOcrProposal, PdfFormFieldValue> {
                [name] = new string('N', 31), [reference] = new string('R', 16)
            });

            Assert.Equal(new string('N', 31), result.Inspect().FormFieldsByName["FullName"].Value);
            Assert.Equal(new string('R', 16), result.Inspect().FormFieldsByName["Reference"].Value);
            Assert.Equal(original, source.ToBytes());
            PdfReadLimitException limit = Assert.Throws<PdfReadLimitException>(() => review.Apply(source,
                new Dictionary<PdfFormOcrProposal, PdfFormFieldValue> {
                    [name] = new string('N', 31),
                    [review.Proposals.Single(proposal => proposal.Field.Name == "SerialCode")] = "ABC123"
                }));
            Assert.Equal(PdfReadLimitKind.OcrArtifacts, limit.Kind);
            Assert.Equal(36, limit.Limit);
            Assert.Equal(37, limit.Actual);
            Assert.Equal(original, source.ToBytes());
        }

        [Theory]
        [InlineData(4, true)]
        [InlineData(5, false)]
        public async Task SharedValuesCountOncePerWidgetPageIncludingPagesWithoutTheirEvidence(int otherLength, bool allowed) {
            string[] objects = {
                "<< /Type /Catalog /Pages 2 0 R /AcroForm 11 0 R >>",
                "<< /Type /Pages /Count 2 /Kids [3 0 R 4 0 R] >>",
                "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources << >> /Annots [6 0 R 7 0 R] >>",
                "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources << >> /Annots [8 0 R 10 0 R] >>",
                "<< /FT /Tx /T (Shared) /V () /Kids [6 0 R 7 0 R 8 0 R] >>",
                "<< /Type /Annot /Subtype /Widget /Parent 5 0 R /P 3 0 R /Rect [20 100 180 120] /F 4 >>",
                "<< /Type /Annot /Subtype /Widget /Parent 5 0 R /P 3 0 R /Rect [20 60 180 80] /F 4 >>",
                "<< /Type /Annot /Subtype /Widget /Parent 5 0 R /P 4 0 R /Rect [20 100 180 120] /F 4 >>",
                "<< /FT /Tx /T (Other) /V () /Kids [10 0 R] >>",
                "<< /Type /Annot /Subtype /Widget /Parent 9 0 R /P 4 0 R /Rect [20 60 180 80] /F 4 >>",
                "<< /Fields [5 0 R 9 0 R] /DA (/F1 10 Tf 0 g) /DR << /Font << /F1 12 0 R >> >> >>",
                "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
                "<< /Title (Shared form widget budgets) >>"
            };
            byte[] original = PdfPageExtractor.Assemble(objects.Select((value, index) => PdfPageExtractor.WrapObject(index + 1,
                Encoding.ASCII.GetBytes(value))).ToList(), 1, 13, PdfFileVersion.Pdf17);
            PdfDocument source = PdfDocument.Load(original);
            IReadOnlyList<PdfFormField> fields = source.Inspect().FormFields;
            Dictionary<int, PdfPageLayoutInfo> layouts = source.GetPageLayouts().ToDictionary(page => page.PageNumber);
            DelegateOcrEngine engine = new DelegateOcrEngine("widget-budget-fixture", (request, _) => Task.FromResult(new OcrResult {
                Spans = fields.SelectMany(field => field.Widgets.Where(widget => widget.PageNumber == request.PageNumber &&
                    (field.Name != "Shared" || request.PageNumber == 1)).Select(widget => {
                        PdfSelectionQuad bounds = layouts[widget.PageNumber!.Value].MapUserSpaceRectangleToVisual(widget.X1, widget.Y1, widget.X2, widget.Y2);
                        return new OcrTextSpan {
                            Text = "x", Confidence = .99, Level = OcrTextSpanLevel.Word, CoordinateUnit = OcrCoordinateUnit.Points,
                            Region = new() { X = bounds.Left + 2, Y = bounds.Top + 2, Width = bounds.Width - 4, Height = bounds.Height - 4 }
                        };
                    })).ToArray()
            }));
            PdfFormOcrReview review = await source.PrepareFormOcrAsync(engine, new() { Dpi = 72, MaxOcrTextCharactersPerPage = 10 });
            PdfFormOcrProposal shared = review.Proposals.Single(proposal => proposal.Field.Name == "Shared");
            Assert.Equal(new[] { 1, 1 }, shared.Evidence.Select(evidence => evidence.PageNumber));
            Dictionary<PdfFormOcrProposal, PdfFormFieldValue> accepted = new Dictionary<PdfFormOcrProposal, PdfFormFieldValue> {
                [shared] = "Shared", [review.Proposals.Single(proposal => proposal.Field.Name == "Other")] = new string('O', otherLength)
            };

            if (allowed) {
                PdfDocument result = review.Apply(source, accepted);
                Assert.Equal("Shared", result.Inspect().FormFieldsByName["Shared"].Value);
                Assert.Equal(new string('O', otherLength), result.Inspect().FormFieldsByName["Other"].Value);
            } else {
                PdfReadLimitException limit = Assert.Throws<PdfReadLimitException>(() => review.Apply(source, accepted));
                Assert.Equal(10, limit.Limit);
                Assert.Equal(11, limit.Actual);
            }
            Assert.Equal(original, source.ToBytes());
        }
    }
}
