using System;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf {
    public partial class PdfComposePageOptionsTests {
        [Fact]
        public void FooterText_RejectsNullConfigurationAndTextSegments() {
            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Header(header => header.Text((string)null!)))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Header(header => header.Text((Action<HeaderTextBuilder>)null!)))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Header(header => header.FirstPageText((string)null!)))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Header(header => header.FirstPageText((Action<HeaderTextBuilder>)null!)))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Header(header => header.EvenPagesText((string)null!)))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Header(header => header.EvenPagesText((Action<HeaderTextBuilder>)null!)))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Footer(footer => footer.Text((string)null!)))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Footer(footer => footer.Text((Action<FooterTextBuilder>)null!)))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Footer(footer => footer.FirstPageText((string)null!)))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Footer(footer => footer.FirstPageText((Action<FooterTextBuilder>)null!)))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Footer(footer => footer.EvenPagesText((string)null!)))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDocument.Create().Compose(c => c.Page(page => page.Footer(footer => footer.EvenPagesText((Action<FooterTextBuilder>)null!)))));

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
        }

        [Fact]
        public void FooterSegments_RejectInvalidExternalState() {
            var nullEntryOptions = new PdfOptions {
                ShowPageNumbers = true,
                FooterSegments = new System.Collections.Generic.List<FooterSegment> { null! }
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
                new FooterSegment(FooterSegmentKind.Text, null));
            Assert.Equal("text", nullTextException.ParamName);
            Assert.Contains("Footer text segments cannot be null.", nullTextException.Message, StringComparison.Ordinal);

            var invalidKindException = Assert.Throws<ArgumentOutOfRangeException>(() =>
                new FooterSegment((FooterSegmentKind)99));
            Assert.Equal("kind", invalidKindException.ParamName);
            Assert.Contains("Footer segments must use a supported segment kind.", invalidKindException.Message, StringComparison.Ordinal);

            var textSegment = new FooterSegment(FooterSegmentKind.Text, string.Empty);
            var pageSegment = new FooterSegment(FooterSegmentKind.PageNumber);
            var totalSegment = new FooterSegment(FooterSegmentKind.TotalPages);

            Assert.Equal(string.Empty, textSegment.Text);
            Assert.Null(pageSegment.Text);
            Assert.Null(totalSegment.Text);
        }

        [Fact]
        public void FooterSegments_SnapshotAssignedAndReadbackLists() {
            var assigned = new System.Collections.Generic.List<FooterSegment> {
                new FooterSegment(FooterSegmentKind.Text, "Page "),
                new FooterSegment(FooterSegmentKind.PageNumber)
            };

            var options = new PdfOptions {
                ShowPageNumbers = true,
                FooterSegments = assigned
            };

            assigned[0] = new FooterSegment(FooterSegmentKind.Text, "Mutated");
            assigned.Add(new FooterSegment(FooterSegmentKind.TotalPages));

            var readback = options.FooterSegments!;
            readback[0] = new FooterSegment(FooterSegmentKind.Text, "Readback mutated");
            readback.Add(new FooterSegment(FooterSegmentKind.TotalPages));

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
