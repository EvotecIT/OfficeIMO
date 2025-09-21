using System;
using System.Linq;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf {
    public class PdfRowComposeTests {
        [Fact]
        public void ColumnRejectsNonPositiveWidth() {
            var doc = PdfDoc.Create();

            Assert.Throws<ArgumentOutOfRangeException>(() =>
                doc.Compose(compose =>
                    compose.Page(page =>
                        page.Content(content =>
                            content.Row(row =>
                                row.Column(0, _ => { }))))));
        }

        [Fact]
        public void ColumnRejectsWidthOverOneHundred() {
            var doc = PdfDoc.Create();

            Assert.Throws<ArgumentOutOfRangeException>(() =>
                doc.Compose(compose =>
                    compose.Page(page =>
                        page.Content(content =>
                            content.Row(row =>
                                row.Column(150, _ => { }))))));
        }

        [Fact]
        public void ColumnRejectsWhenTotalWouldExceedOneHundred() {
            var doc = PdfDoc.Create();

            Assert.Throws<InvalidOperationException>(() =>
                doc.Compose(compose =>
                    compose.Page(page =>
                        page.Content(content =>
                            content.Row(row => {
                                row.Column(60, _ => { });
                                row.Column(50, _ => { });
                            })))));
        }

        [Fact]
        public void ColumnsAreNormalizedWhenTotalLessThanOneHundred() {
            var doc = PdfDoc.Create();

            doc.Compose(compose =>
                compose.Page(page =>
                    page.Content(content =>
                        content.Row(row => {
                            row.Column(30, _ => { });
                            row.Column(20, _ => { });
                        }))));

            var page = Assert.IsType<PageBlock>(Assert.Single(doc.Blocks));
            var row = Assert.IsType<RowBlock>(Assert.Single(page.Blocks));
            Assert.Equal(2, row.Columns.Count);

            var widths = row.Columns.Select(c => c.WidthPercent).ToArray();
            Assert.InRange(widths[0], 59.9, 60.1);
            Assert.InRange(widths[1], 39.9, 40.1);
            var total = widths.Sum();
            Assert.InRange(total, 99.99, 100.01);
        }
    }
}
