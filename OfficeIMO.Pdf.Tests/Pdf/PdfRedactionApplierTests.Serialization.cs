using OfficeIMO.Pdf;
using System;
using System.Linq;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests.Pdf {
    public partial class PdfRedactionApplierTests {
        [Theory]
        [InlineData("Rectangle")]
        [InlineData("Polygon")]
        [InlineData("Freehand")]
        public void RedactionMarksPreserveFractionalReviewedGeometry(string shape) {
            byte[] bytes = BuildTextContentRedactionSource(
                "75 600 1 1 re f BT /F1 12 Tf 72 720 Td (Public summary) Tj ET");
            const double left = 70.2578125D, bottom = 598.30078125D, width = 150.388671875D, height = 26.6484375D;
            PdfRedactionArea[] areas = shape switch {
                "Polygon" => PdfRedactionRegion.Polygon(1, new[] {
                    new PdfRedactionPoint(left, bottom), new PdfRedactionPoint(left + width, bottom),
                    new PdfRedactionPoint(left + width, bottom + height), new PdfRedactionPoint(left, bottom + height)
                }).Areas.ToArray(),
                "Freehand" => PdfRedactionRegion.Freehand(1, new[] {
                    new PdfRedactionPoint(left, 600.2578125D), new PdfRedactionPoint(left + width, 600.2578125D)
                }, 12.6484375D).Areas.ToArray(),
                _ => new[] { new PdfRedactionArea(1, left, bottom, width, height) }
            };
            PdfDocument source = PdfDocument.Load(bytes);
            PdfRedactionPlan plan = source.Redactions.Plan(areas);
            Assert.Contains(plan.Matches, match => match.Kind == PdfRedactionMatchKind.VectorPath);

            PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(plan);

            Assert.True(result.Evidence.IsVerified, result.Evidence.Summary);
            Assert.Contains("Public summary", result.ToDocument().Read().Text, StringComparison.Ordinal);
            Assert.Equal(bytes, source.ToBytes());
        }

        [Fact]
        public void PreciseRedactionPreservesImportedFontDescriptorNumbers() {
            byte[] sourceBytes = BuildPdf(new[] {
                "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
                "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>\nendobj",
                "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R /Resources << /Font << /F1 4 0 R >> >> >>\nendobj",
                "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding /FontDescriptor 6 0 R >>\nendobj",
                BuildStreamObject(5, Encoding.ASCII.GetBytes("BT /F1 18 Tf 72 720 Td (Alpha secret Omega) Tj ET")),
                "6 0 obj\n<< /Type /FontDescriptor /FontName /Helvetica /Flags 32 /FontBBox [-489.2578 -258.30079 1147.9492 1014.64846] /Ascent 1014.64846 /Descent -258.30079 /CapHeight 700 /ItalicAngle 0 /StemV 80 >>\nendobj"
            }, rootObjectNumber: 1);
            PdfDocument source = PdfDocument.Load(sourceBytes);
            PdfRedactionPlan plan = source.Redactions.Search(PreciseOptions().AddLiteral("secret"));
            Assert.True(plan.IsReviewable, DescribeFindings(plan));

            PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(plan);

            Assert.True(result.Evidence.IsVerified, result.Evidence.Summary);
            Assert.Equal(ReadFontBounds(sourceBytes), ReadFontBounds(result.Pdf));
            Assert.DoesNotContain("secret", result.ToDocument().Read().Text, StringComparison.Ordinal);
            Assert.Contains("Alpha", result.ToDocument().Read().Text, StringComparison.Ordinal);
            Assert.Contains("Omega", result.ToDocument().Read().Text, StringComparison.Ordinal);
            Assert.Equal(sourceBytes, source.ToBytes());
        }

        private static double[] ReadFontBounds(byte[] pdf) {
            PdfReadDocument document = PdfReadDocument.Open(pdf);
            PdfDictionary descriptor = document.Objects.Values.Select(static item => item.Value)
                .OfType<PdfDictionary>().Single(static item => item.Get<PdfName>("Type")?.Name == "FontDescriptor");
            return descriptor.Get<PdfArray>("FontBBox")!.Items.Cast<PdfNumber>().Select(static number => number.Value).ToArray();
        }
    }
}