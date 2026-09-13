using System;
using System.Linq;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class DrawingTablePaginationTests {
        [Fact]
        public void MeasuredRowsRepeatHeadersAndPreserveOrderAtExactBoundaries() {
            var pages = OfficeTablePagination.Paginate(new[] { 20d, 50d, 40d, 30d, 10d }, 80, 10);
            Assert.Equal(new[] { 0, 2, 4 }, pages.Select(p => p.RowOffset));
            Assert.Equal(new[] { 2, 2, 1 }, pages.Select(p => p.RowCount));
            Assert.Equal(new[] { 80d, 80d, 20d }, pages.Select(p => p.Height));
            Assert.Equal(5, pages.Sum(p => p.RowCount));
        }

        [Fact]
        public void EmptyTablesRetainHeaderAndHonorCancellation() {
            var page = Assert.Single(OfficeTablePagination.Paginate(Array.Empty<double>(), 80, 20));
            Assert.Equal(0, page.RowCount); Assert.Equal(20, page.Height);
            Assert.Throws<OperationCanceledException>(() => OfficeTablePagination.Paginate(Array.Empty<double>(), 80, 20, cancellationToken: new CancellationToken(true)));
        }

        [Fact]
        public void UnfitRowsAndPageBudgetsFailWithoutPartialResults() {
            Assert.Throws<InvalidOperationException>(() => OfficeTablePagination.Paginate(new[] { 71d }, 80, 10));
            Assert.Throws<InvalidOperationException>(() => OfficeTablePagination.Paginate(new[] { 70d, 1d }, 80, 10, 1));
            Assert.Single(OfficeTablePagination.Paginate(new[] { 70d }, 80, 10, 1));
        }

        [Theory]
        [InlineData(double.NaN)]
        [InlineData(double.PositiveInfinity)]
        [InlineData(0d)]
        [InlineData(-1d)]
        public void InvalidRowsCannotProduceInvalidGeometry(double height) =>
            Assert.Throws<ArgumentOutOfRangeException>(() => OfficeTablePagination.Paginate(new[] { height }, 80, 10));
    }
}
