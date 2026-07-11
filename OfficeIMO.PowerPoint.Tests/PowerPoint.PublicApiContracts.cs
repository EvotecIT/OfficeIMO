using System;
using System.IO;
using System.Linq;
using System.Reflection;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointPublicApiContracts {
        [Fact]
        public void PublicSurfaceHasOneLifecycleTemplateAndCompositionPath() {
            Assembly assembly = typeof(PowerPointPresentation).Assembly;
            string[] exportedTypes = assembly.GetExportedTypes().Select(type => type.FullName!).ToArray();
            Assert.DoesNotContain("OfficeIMO.PowerPoint.PowerPointDeckComposer", exportedTypes);
            Assert.DoesNotContain("OfficeIMO.PowerPoint.PowerPointSlideComposer", exportedTypes);
            Assert.DoesNotContain("OfficeIMO.PowerPoint.PowerPointDesignExtensions", exportedTypes);
            Assert.DoesNotContain("OfficeIMO.PowerPoint.PowerPointChartData", exportedTypes);
            Assert.DoesNotContain("OfficeIMO.PowerPoint.PowerPointScatterChartData", exportedTypes);
            Assert.DoesNotContain("OfficeIMO.PowerPoint.PowerPointChartSnapshot", exportedTypes);
            Assert.DoesNotContain(exportedTypes, name => name.StartsWith("OfficeIMO.PowerPoint.Fluent.",
                StringComparison.Ordinal));

            MethodInfo[] presentationMethods = typeof(PowerPointPresentation).GetMethods(BindingFlags.Public |
                BindingFlags.Instance | BindingFlags.Static | BindingFlags.DeclaredOnly);
            Assert.Single(presentationMethods, method => method.Name == nameof(PowerPointPresentation.Compose));
            Assert.DoesNotContain(presentationMethods, method => method.Name == "OpenRead");
            Assert.DoesNotContain(presentationMethods, method => method.Name == "CreateFromTemplate");
            Assert.DoesNotContain(presentationMethods, method => method.Name == "InspectTemplate");
            Assert.DoesNotContain(presentationMethods, method => method.Name == "Preflight");
            Assert.DoesNotContain(presentationMethods, method => method.Name == "SaveWithPreflight");
            Assert.DoesNotContain(presentationMethods, method => method.Name == "CreateVisualProofReport");

            Assert.Contains(presentationMethods, method => method.Name == nameof(PowerPointPresentation.Inspect));
            Assert.Contains(presentationMethods, method => method.Name == nameof(PowerPointPresentation.InspectPreflight));
            Assert.Contains(presentationMethods, method => method.Name == nameof(PowerPointPresentation.InspectVisuals));
            Assert.NotNull(typeof(PowerPointTemplate).GetMethod(nameof(PowerPointTemplate.Inspect),
                new[] { typeof(string) }));
            Assert.NotNull(typeof(PowerPointTemplate).GetMethod(nameof(PowerPointTemplate.CreatePresentation)));
        }

        [Fact]
        public void SemanticCompositionProducesEditableRoundTripAndUnifiedInspection() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                PowerPointDeckPlan plan = new PowerPointDeckPlan()
                    .AddSection("Operating model", "One semantic plan and one renderer")
                    .AddCardGrid("Ownership", "Concrete objects remain editable", new[] {
                        new PowerPointCardContent("Lifecycle", new[] { "Create", "Open", "Save" }),
                        new PowerPointCardContent("Editing", new[] { "Slides", "Shapes", "Content" })
                    });
                PowerPointDeckDesign design = PowerPointDeckDesign.FromBrand("#008C95", "public-api-contract");

                using (PowerPointPresentation presentation = PowerPointPresentation.Create(path)) {
                    PowerPointCompositionResult result = presentation.Compose(plan,
                        PowerPointCompositionOptions.FromDesign(design));
                    Assert.Equal(2, result.Slides.Count);
                    Assert.Same(design, result.Design);
                    Assert.True(result.Preflight.IsSuccessful);

                    result.Slides[1].AddTextBoxPoints("Edited after composition", 40, 430, 260, 30);
                    PowerPointInspectionReport inspection = presentation.Inspect();
                    Assert.Empty(inspection.PackageErrors);
                    Assert.NotNull(inspection.Preflight);
                    Assert.NotNull(inspection.Accessibility);
                    presentation.Save();
                }

                using PowerPointPresentation reopened = PowerPointPresentation.Open(path,
                    PowerPointOpenMode.ReadOnly);
                Assert.Contains(reopened.Slides.SelectMany(slide => slide.TextBoxes),
                    textBox => textBox.Text == "Edited after composition");
                Assert.Empty(reopened.ValidateDocument());
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void StreamOptionsMakePersistenceAndReadOnlyIntentExplicit() {
            using var stream = new MemoryStream();
            using (PowerPointPresentation presentation = PowerPointPresentation.Create(stream,
                       new PowerPointStreamCreateOptions { AutoSave = true })) {
                Assert.Empty(presentation.Slides);
                presentation.AddSlide().AddTitle("Stream lifecycle");
                Assert.Single(presentation.Slides);
            }

            using PowerPointPresentation reopened = PowerPointPresentation.Open(stream,
                new PowerPointStreamOpenOptions { Mode = PowerPointOpenMode.ReadOnly });
            Assert.Equal("Stream lifecycle", reopened.Slides[0].TextBoxes.First().Text);
        }

        [Fact]
        public void CompositionValidatesBeforeApplyingThemeOrAddingSlides() {
            var missingImage = new PowerPointImageAsset(
                Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".png"), "Missing screenshot");
            PowerPointDeckPlan invalidPlan = new PowerPointDeckPlan()
                .AddScreenshotStory("Invalid proof", null, missingImage);
            PowerPointDeckDesign design = PowerPointDeckDesign.FromBrand("#D93025", "invalid-plan-theme");
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream,
                new PowerPointStreamCreateOptions { AutoSave = false });
            string originalThemeName = presentation.ThemeName;
            string? originalAccent = presentation.GetThemeColor(PowerPointThemeColor.Accent1);

            Assert.Throws<PowerPointDeckPlanValidationException>(() =>
                presentation.Compose(invalidPlan, PowerPointCompositionOptions.FromDesign(design)));

            Assert.Equal(originalThemeName, presentation.ThemeName);
            Assert.Equal(originalAccent, presentation.GetThemeColor(PowerPointThemeColor.Accent1));
            Assert.Empty(presentation.Slides);
        }
    }
}
