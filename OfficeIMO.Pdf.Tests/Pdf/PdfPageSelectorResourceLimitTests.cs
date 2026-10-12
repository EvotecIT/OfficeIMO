using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf {
    public class PdfPageSelectorResourceLimitTests {
        [Theory]
        [InlineData("all,!all")]
        [InlineData("!odd,!even")]
        [InlineData("1,!1..last")]
        [InlineData("all,!1..last-1,!last")]
        public void BoundedResolutionRejectsFullyExcludedSelectionsAtTheIntegerPageCountLimit(string expression) {
            InvalidOperationException error = Assert.Throws<InvalidOperationException>(() =>
                PdfPageSelector.Parse(expression).Resolve(int.MaxValue, 2));

            Assert.Equal("Page selector resolved to an empty page set.", error.Message);
        }

        [Theory]
        [InlineData("all,!1..last-1", int.MaxValue)]
        [InlineData("last..1,!1..last-1", int.MaxValue)]
        [InlineData("!2..last", 1)]
        [InlineData("odd,!1..last-2", int.MaxValue)]
        [InlineData("even,!1..last-3", int.MaxValue - 1)]
        public void BoundedResolutionRetainsOnlyTheSurvivorOfLargeExcludedRuns(string expression, int expected) {
            Assert.Equal(new[] { expected }, PdfPageSelector.Parse(expression).Resolve(int.MaxValue, 1));
        }

        [Theory]
        [InlineData("all,!1..last-3")]
        [InlineData("last..1,!1..last-3")]
        [InlineData("!1..last-3")]
        public void BoundedResolutionEnforcesTheAcceptedPageLimitAfterSkippingExcludedRuns(string expression) {
            InvalidOperationException error = Assert.Throws<InvalidOperationException>(() =>
                PdfPageSelector.Parse(expression).Resolve(int.MaxValue, 2));

            Assert.Equal("Page selection exceeds the configured page limit.", error.Message);
        }

        [Fact]
        public void ResolveMergesReversedOverlappingAndAdjacentExclusionsWithoutChangingOrderOrDuplicates() {
            PdfPageSelector selector = PdfPageSelector.Parse("last..1,1..last,last,3,!6..5,!4..2,!3");

            Assert.Equal(new[] { 8, 7, 1, 1, 7, 8, 8 }, selector.Resolve(8, 7));
            Assert.Throws<InvalidOperationException>(() => selector.Resolve(8, 6));
        }

        [Theory]
        [InlineData("last..1,odd,even,1,!2..4,!even", new[] { 9, 7, 5, 1, 1, 5, 7, 9, 1 })]
        [InlineData("last..1,odd,even,2,!2..4,!odd", new[] { 8, 6, 6, 8 })]
        public void ResolveCombinesParityAndRangeExclusionsAcrossOrderedInclusions(string expression, int[] expected) {
            Assert.Equal(expected, PdfPageSelector.Parse(expression).Resolve(9, expected.Length));
        }

        [Theory]
        [InlineData("all,!all,!last-9")]
        [InlineData("last-9,!all")]
        [InlineData("!odd,!even,!last-9")]
        [InlineData("last-9,!odd,!even")]
        [InlineData("1..last-9,!1..last")]
        public void ResolveStillValidatesEndpointsWhenTheSelectionIsEntirelyExcluded(string expression) {
            Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageSelector.Parse(expression).Resolve(9, 1));
        }
    }
}
