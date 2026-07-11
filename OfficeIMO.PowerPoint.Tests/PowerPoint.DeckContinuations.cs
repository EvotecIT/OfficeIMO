using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointDeckContinuationTests {
        [Fact]
        public void WithContinuations_SplitsDenseSemanticContentWithoutDroppingItems() {
            var plan = new PowerPointDeckPlan()
                .AddProcess("Delivery", "Every step remains editable",
                    Enumerable.Range(1, 12)
                        .Select(index => new PowerPointProcessStep("Step " + index, "Body " + index)))
                .AddCardGrid("Portfolio", "All cards continue",
                    Enumerable.Range(1, 13)
                        .Select(index => new PowerPointCardContent("Card " + index)))
                .AddCapability("Capabilities", "All sections continue",
                    Enumerable.Range(1, 9)
                        .Select(index => new PowerPointCapabilitySection("Capability " + index)));

            PowerPointDeckPlan expanded = plan.WithContinuations();
            PowerPointDeckPlanSlideSummary[] summaries = expanded.DescribeSlides().ToArray();

            Assert.Equal(9, summaries.Length);
            Assert.Equal(new[] { 5, 5, 2 }, summaries
                .Where(summary => summary.Kind == PowerPointDeckPlanSlideKind.Process)
                .Select(summary => summary.ContentItemCount));
            Assert.Equal(new[] { 6, 6, 1 }, summaries
                .Where(summary => summary.Kind == PowerPointDeckPlanSlideKind.CardGrid)
                .Select(summary => summary.ContentItemCount));
            Assert.Equal(new[] { 4, 4, 1 }, summaries
                .Where(summary => summary.Kind == PowerPointDeckPlanSlideKind.Capability)
                .Select(summary => summary.ContentItemCount));
            Assert.Equal("Delivery", summaries[0].Title);
            Assert.Equal("Delivery (continued 2/3)", summaries[1].Title);
            Assert.Empty(expanded.ValidateSlides().Where(diagnostic =>
                diagnostic.Severity == PowerPointDeckPlanDiagnosticSeverity.Error));
        }

        [Fact]
        public void Compose_RendersAllProcessStepsAcrossValidContinuationSlides() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            var plan = new PowerPointDeckPlan().AddProcess("Twelve steps", null,
                Enumerable.Range(1, 12)
                    .Select(index => new PowerPointProcessStep("Step " + index, "Body " + index)));

            PowerPointCompositionResult result = presentation.Compose(plan,
                PowerPointCompositionOptions.FromDesign(
                    PowerPointDeckDesign.FromBrand("#2463EB", "continuation-test")));

            Assert.Equal(3, result.Slides.Count);
            string allText = string.Join("\n", result.Slides.SelectMany(slide => slide.TextBoxes).Select(box => box.Text));
            for (int index = 1; index <= 12; index++) {
                Assert.Contains("Step " + index, allText);
            }
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void Compose_ReturnsSharedPreflightContract() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            var plan = new PowerPointDeckPlan().AddSection("Measured deck", "Designer report");

            PowerPointCompositionOptions options = PowerPointCompositionOptions.FromDesign(
                PowerPointDeckDesign.FromBrand("#2463EB", "report-test"));
            options.Preflight = new PowerPointDeckPreflightOptions {
                DetectShapeCollisions = false,
                IncludeVisualSnapshotDiagnostics = false
            };
            PowerPointCompositionResult result = presentation.Compose(plan, options);

            Assert.Single(result.Slides);
            Assert.Equal(presentation.Slides.Count, result.Preflight.SlideCount);
            Assert.Contains("\"schemaVersion\": 1", result.Preflight.ToJson());
        }
    }
}
