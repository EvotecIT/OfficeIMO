using System;
using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointDesignerTextLayout {
        [Theory]
        [InlineData(PowerPointSectionLayoutVariant.GeometricCover)]
        [InlineData(PowerPointSectionLayoutVariant.EditorialRail)]
        [InlineData(PowerPointSectionLayoutVariant.Poster)]
        public void SectionUnderline_ClearsTitleAndSubtitle(PowerPointSectionLayoutVariant variant) {
            WithSavedDeck(presentation => {
                presentation.AddDesignerSectionSlide("Workplace support plan", "Demand review and service pilot",
                    options: new PowerPointDesignerSlideOptions {
                        SectionVariant = variant,
                        TitleAccentStyle = PowerPointTitleAccentStyle.Underline,
                        FooterLeft = "30 Sep 2026", FooterRight = "Synthetic data"
                    });
            }, presentation => {
                PowerPointSlide slide = presentation.Slides[0];
                PowerPointTextBox title = slide.TextBoxes.Single(box => box.Text == "Workplace support plan");
                PowerPointTextBox subtitle = slide.TextBoxes.Single(box => box.Text == "Demand review and service pilot");
                PowerPointShape underline = slide.GetShape("Section Title Accent Underline")!;
                Assert.True(underline.Top >= title.Bottom, "The accent crosses the title's text box.");
                Assert.True(underline.Bottom <= subtitle.Top, "The accent crosses the subtitle's text box.");
                foreach (PowerPointTextBox footer in slide.TextBoxes.Where(box =>
                    box.Text == "30 Sep 2026" || box.Text == "Synthetic data")) {
                    Assert.True(footer.HeightCm * 72 / 2.54 >= footer.FontSize * 1.2,
                        "The footer box does not contain a complete line at its declared font size.");
                    Assert.True(footer.Bottom <= presentation.SlideSize.HeightEmus);
                }
            });
        }

        [Theory]
        [InlineData(PowerPointCaseStudyLayoutVariant.EditorialSplit)]
        [InlineData(PowerPointCaseStudyLayoutVariant.VisualBand)]
        [InlineData(PowerPointCaseStudyLayoutVariant.VisualHero)]
        public void CompactMetricValues_FitTheirVerticalSpace(PowerPointCaseStudyLayoutVariant variant) {
            WithSavedDeck(presentation => {
                presentation.AddDesignerCaseStudySlide("Service demand",
                    new[] { new PowerPointCaseStudySection("Observation", "Review the quarterly intake.") },
                    new[] { new PowerPointMetric("600", "people"), new PowerPointMetric("4", "locations"),
                        new PowerPointMetric("1,080", "tickets") },
                    options: new PowerPointCaseStudySlideOptions { Variant = variant, FooterLeft = "30 Sep 2026" });
            }, presentation => {
                PowerPointTextBox[] values = presentation.Slides[0].TextBoxes
                    .Where(box => new[] { "600", "4", "1,080" }.Contains(box.Text)).ToArray();
                Assert.Equal(3, values.Length);
                foreach (PowerPointTextBox value in values) {
                    Assert.True(value.HeightCm * 72 / 2.54 >= value.FontSize * 1.2,
                        "The metric's font exceeds its available vertical space.");
                }
                if (variant == PowerPointCaseStudyLayoutVariant.VisualBand) {
                    PowerPointSlide slide = presentation.Slides[0];
                    foreach (string text in new[] { "Observation", "OfficeIMO", "Project portfolio" }) {
                        PowerPointTextBox label = slide.TextBoxes.Single(box => box.Text == text);
                        Assert.True(label.HeightCm * 72 / 2.54 >= label.FontSize * 1.2);
                    }
                    PowerPointTextBox heading = slide.TextBoxes.Single(box => box.Text == "Observation");
                    Assert.True(slide.GetShape("Case Study Client Rule")!.Top >= heading.Bottom);
                }
            });
        }

        [Theory]
        [InlineData(PowerPointMetricStripVariant.SolidBand)]
        [InlineData(PowerPointMetricStripVariant.SeparatedTiles)]
        [InlineData(PowerPointMetricStripVariant.Underlined)]
        public void ComposerMetricValues_FitCompactStrips(PowerPointMetricStripVariant variant) {
            WithSavedDeck(presentation => {
                presentation.ComposeDesignerSlide(composer => composer.AddMetricStrip(
                    new[] { new PowerPointMetric("600", "people"), new PowerPointMetric("4", "locations") },
                    PowerPointLayoutBox.FromCentimeters(2, 5, 18, 1.25), variant));
            }, presentation => {
                PowerPointTextBox[] values = presentation.Slides[0].TextBoxes
                    .Where(box => box.Text == "600" || box.Text == "4").ToArray();
                Assert.Equal(2, values.Length);
                Assert.All(values, value => Assert.True(value.HeightCm * 72 / 2.54 >= value.FontSize * 1.2));
            });
        }

        private static void WithSavedDeck(Action<PowerPointDocument> build, Action<PowerPointDocument> inspect) {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".pptx");
            try {
                using (PowerPointDocument presentation = PowerPointDocument.Create(path)) {
                    presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);
                    build(presentation);
                    presentation.Save();
                }
                using (PowerPointDocument presentation = PowerPointDocument.Load(path)) {
                    Assert.Empty(presentation.ValidateDocument());
                    inspect(presentation);
                }
            } finally {
                if (File.Exists(path)) { File.Delete(path); }
            }
        }
    }
}
