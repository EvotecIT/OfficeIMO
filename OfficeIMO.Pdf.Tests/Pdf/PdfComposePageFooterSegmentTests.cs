using System;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf {
    public partial class PdfDocumentBuilderPageOptionsTests {
        [Fact]
        public void FooterText_RejectsNullConfigurationAndTextSegments() {
            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Header(header => header.Text((string)null!)))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Header(header => header.Text((Action<PdfHeaderTextBuilder>)null!)))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Header(header => header.FirstPageText((string)null!)))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Header(header => header.FirstPageText((Action<PdfHeaderTextBuilder>)null!)))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Header(header => header.EvenPagesText((string)null!)))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Header(header => header.EvenPagesText((Action<PdfHeaderTextBuilder>)null!)))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Footer(footer => footer.Text((string)null!)))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Footer(footer => footer.Text((Action<PdfFooterTextBuilder>)null!)))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Footer(footer => footer.FirstPageText((string)null!)))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Footer(footer => footer.FirstPageText((Action<PdfFooterTextBuilder>)null!)))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Footer(footer => footer.EvenPagesText((string)null!)))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Footer(footer => footer.EvenPagesText((Action<PdfFooterTextBuilder>)null!)))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Footer(footer => footer.Text(text => text.Text(null!))))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Header(header => header.Text(text => text.Text(null!))))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Header(header => header.FirstPageText(text => text.Text(null!))))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Header(header => header.EvenPagesText(text => text.Text(null!))))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Footer(footer => footer.FirstPageText(text => text.Text(null!))))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Footer(footer => footer.EvenPagesText(text => text.Text(null!))))));

            Assert.Throws<ArgumentException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Header(header => header.Zones(null, null, null)))));

            Assert.Throws<ArgumentException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Footer(footer => footer.Zones(null, null, null)))));

            Assert.Throws<ArgumentException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Header(header => header.FirstPageZones(null, null, null)))));

            Assert.Throws<ArgumentException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Footer(footer => footer.EvenPagesZones(null, null, null)))));

            Assert.Throws<ArgumentException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Header(header => header.StyledZones(null, null, null)))));

            Assert.Throws<ArgumentException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Footer(footer => footer.FirstPageStyledZones(null, null, null)))));
        }

        [Fact]
        public void FooterSegments_RejectInvalidExternalState() {
            var nullEntryOptions = new PdfCreateOptions {
                ShowPageNumbers = true,
                FooterSegments = new System.Collections.Generic.List<PdfFooterSegment> { null! }
            };

            var nullEntryException = Assert.Throws<ArgumentException>(() =>
                PdfDocument.Create(nullEntryOptions)
                    .Paragraph(p => p.Text("Invalid footer segment"))
                    .ToBytes());
            Assert.Contains("footer segments", nullEntryException.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void FooterSegment_RejectsInvalidIntrinsicStateAtConstruction() {
            var nullTextException = Assert.Throws<ArgumentNullException>(() =>
                new PdfFooterSegment(PdfFooterSegmentKind.Text, null));
            Assert.Equal("text", nullTextException.ParamName);
            Assert.Contains("Footer text segments cannot be null.", nullTextException.Message, StringComparison.Ordinal);

            var invalidKindException = Assert.Throws<ArgumentOutOfRangeException>(() =>
                new PdfFooterSegment((PdfFooterSegmentKind)99));
            Assert.Equal("kind", invalidKindException.ParamName);
            Assert.Contains("Footer segments must use a supported segment kind.", invalidKindException.Message, StringComparison.Ordinal);

            var textSegment = new PdfFooterSegment(PdfFooterSegmentKind.Text, string.Empty);
            var pageSegment = new PdfFooterSegment(PdfFooterSegmentKind.PageNumber);
            var totalSegment = new PdfFooterSegment(PdfFooterSegmentKind.TotalPages);

            Assert.Equal(string.Empty, textSegment.Text);
            Assert.Null(pageSegment.Text);
            Assert.Null(totalSegment.Text);
        }

        [Fact]
        public void FooterSegments_SnapshotAssignedAndReadbackLists() {
            var assigned = new System.Collections.Generic.List<PdfFooterSegment> {
                new PdfFooterSegment(PdfFooterSegmentKind.Text, "Page "),
                new PdfFooterSegment(PdfFooterSegmentKind.PageNumber)
            };

            var options = new PdfCreateOptions {
                ShowPageNumbers = true,
                FooterSegments = assigned
            };

            assigned[0] = new PdfFooterSegment(PdfFooterSegmentKind.Text, "Mutated");
            assigned.Add(new PdfFooterSegment(PdfFooterSegmentKind.TotalPages));

            var readback = options.FooterSegments!;
            readback[0] = new PdfFooterSegment(PdfFooterSegmentKind.Text, "Readback mutated");
            readback.Add(new PdfFooterSegment(PdfFooterSegmentKind.TotalPages));

            Assert.Equal(2, options.FooterSegments!.Count);
            Assert.Equal("Page ", options.FooterSegments![0].Text);

            var doc = PdfDocument.Create(options)
                .Paragraph(p => p.Text("Footer segment snapshot"));

            string pdfText = Normalize(PdfPigDocument.Open(new MemoryStream(doc.ToBytes())).GetPage(1).Text);
            Assert.Contains("Page1", pdfText, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Mutated", pdfText, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Readbackmutated", pdfText, StringComparison.OrdinalIgnoreCase);
        }

    }
}
