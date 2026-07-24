using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests {
    public class PowerPointTemplateWorkflowTests {
        [Fact]
        public void InspectTemplate_MapsLayoutPlaceholdersThemeAndBrandTokens() {
            string templatePath = CreateTemplatePresentation();
            string targetPath = TempPath(".pptx");
            try {
                PowerPointTemplateInventory inventory = PowerPointTemplate.Inspect(templatePath);

                PowerPointTemplateMasterInfo master = Assert.Single(inventory.Masters);
                Assert.Equal("Corporate Test Theme", master.ThemeName);
                Assert.Equal("0B7FAB", master.ThemeColors[PowerPointThemeColor.Accent1]);
                Assert.Equal("Aptos Display", master.ThemeFonts.MajorLatin);
                Assert.NotEmpty(master.Layouts);
                Assert.Equal(3, inventory.SourceSlideCount);

                PowerPointTemplateLayoutInfo layout = inventory.ResolveLayout(master.Layouts[0].Name);
                Assert.True(layout.SafeArea.Width > 0);
                PowerPointTemplatePlaceholderInfo body = layout.ResolvePlaceholder("Executive Summary Body");
                Assert.Equal(PowerPointTemplatePlaceholderRole.Body, body.Role);
                Assert.True(body.Bounds.HasValue);
                PowerPointTemplateResolutionException ambiguous = Assert.Throws<PowerPointTemplateResolutionException>(
                    () => layout.ResolvePlaceholder(PowerPointTemplatePlaceholderRole.Body));
                Assert.Equal("Template.PlaceholderAmbiguous", ambiguous.Code);
                Assert.True(ambiguous.Candidates.Count >= 2);
                Assert.Contains("CONFIDENTIAL", inventory.FooterContents);

                PowerPointDesignBrief brief = inventory.CreateDesignBrief("template-brand", "executive review");
                Assert.Equal("0B7FAB", brief.AccentColor);
                Assert.Equal("Aptos Display", brief.HeadingFontName);
                Assert.Equal("Aptos", brief.BodyFontName);
                Assert.Equal("CONFIDENTIAL", brief.FooterLeft);

                using PowerPointPresentation target = PowerPointPresentation.Create(targetPath);
                inventory.ApplyBrandTo(target);
                Assert.Equal("Corporate Test Theme", target.ThemeName);
                Assert.Equal("0B7FAB", target.GetThemeColor(PowerPointThemeColor.Accent1));
            } finally {
                Delete(templatePath, targetPath);
            }
        }

        [Fact]
        public void InspectTemplate_InventoriesPicturesNestedInsideGroups() {
            string templatePath = CreateTemplatePresentation();
            try {
                using (PresentationDocument document = PresentationDocument.Open(templatePath, true)) {
                    SlideLayoutPart layoutPart = document.PresentationPart!.SlideMasterParts.First()
                        .SlideLayoutParts.First();
                    ShapeTree tree = layoutPart.SlideLayout.CommonSlideData!.ShapeTree!;
                    tree.Append(new GroupShape(
                        new NonVisualGroupShapeProperties(
                            new NonVisualDrawingProperties { Id = 900U, Name = "Brand Group" },
                            new NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()),
                        new GroupShapeProperties(new A.TransformGroup(
                            new A.Offset { X = 2000000L, Y = 1000000L },
                            new A.Extents { Cx = 2000000L, Cy = 1000000L },
                            new A.ChildOffset { X = 0L, Y = 0L },
                            new A.ChildExtents { Cx = 1000000L, Cy = 500000L })),
                        new DocumentFormat.OpenXml.Presentation.Picture(
                            new NonVisualPictureProperties(
                                new NonVisualDrawingProperties { Id = 901U, Name = "Grouped Logo" },
                                new NonVisualPictureDrawingProperties(),
                                new ApplicationNonVisualDrawingProperties()),
                            new BlipFill(new A.Blip(), new A.Stretch(new A.FillRectangle())),
                            new ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset { X = 100000L, Y = 50000L },
                                    new A.Extents { Cx = 200000L, Cy = 100000L }),
                                new A.PresetGeometry { Preset = A.ShapeTypeValues.Rectangle }))));
                    tree.Append(new DocumentFormat.OpenXml.Presentation.Picture(
                        new NonVisualPictureProperties(
                            new NonVisualDrawingProperties { Id = 912U, Name = "Long Boundary Logo" },
                            new NonVisualPictureDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()),
                        new BlipFill(new A.Blip(), new A.Stretch(new A.FillRectangle())),
                        new ShapeProperties(
                            new A.Transform2D(
                                new A.Offset { X = long.MaxValue, Y = 0L },
                                new A.Extents { Cx = 1L, Cy = 1L }),
                            new A.PresetGeometry { Preset = A.ShapeTypeValues.Rectangle })));
                    layoutPart.SlideLayout.Save();
                }

                PowerPointTemplateInventory inventory = PowerPointTemplate.Inspect(templatePath);
                PowerPointTemplateAssetInfo asset = Assert.Single(inventory.Assets,
                    candidate => candidate.Name == "Grouped Logo" &&
                                 candidate.Kind == PowerPointTemplateAssetKind.Logo);
                PowerPointLayoutBox bounds = Assert.IsType<PowerPointLayoutBox>(asset.Bounds);
                Assert.Equal(2200000L, bounds.Left);
                Assert.Equal(1100000L, bounds.Top);
                Assert.Equal(400000L, bounds.Width);
                Assert.Equal(200000L, bounds.Height);
                Assert.Null(Assert.Single(
                    inventory.Assets,
                    candidate => candidate.Name == "Long Boundary Logo").Bounds);
            } finally {
                Delete(templatePath);
            }
        }

        [Fact]
        public void InspectTemplate_DropsNonFiniteGroupedAssetBoundsInsteadOfOverflowing() {
            string templatePath = CreateTemplatePresentation();
            try {
                using (PresentationDocument document = PresentationDocument.Open(templatePath, true)) {
                    SlideLayoutPart layoutPart = document.PresentationPart!.SlideMasterParts.First()
                        .SlideLayoutParts.First();
                    ShapeTree tree = layoutPart.SlideLayout.CommonSlideData!.ShapeTree!;
                    tree.Append(new GroupShape(
                        new NonVisualGroupShapeProperties(
                            new NonVisualDrawingProperties { Id = 910U, Name = "Extreme Group" },
                            new NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()),
                        new GroupShapeProperties(new A.TransformGroup(
                            new A.Offset { X = long.MaxValue, Y = long.MaxValue },
                            new A.Extents { Cx = long.MaxValue, Cy = long.MaxValue },
                            new A.ChildOffset { X = 0L, Y = 0L },
                            new A.ChildExtents { Cx = 1L, Cy = 1L })),
                        new DocumentFormat.OpenXml.Presentation.Picture(
                            new NonVisualPictureProperties(
                                new NonVisualDrawingProperties { Id = 911U, Name = "Extreme Logo" },
                                new NonVisualPictureDrawingProperties(),
                                new ApplicationNonVisualDrawingProperties()),
                            new BlipFill(new A.Blip(), new A.Stretch(new A.FillRectangle())),
                            new ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset { X = long.MaxValue, Y = long.MaxValue },
                                    new A.Extents { Cx = long.MaxValue, Cy = long.MaxValue }),
                                new A.PresetGeometry { Preset = A.ShapeTypeValues.Rectangle }))));
                    layoutPart.SlideLayout.Save();
                }

                PowerPointTemplateInventory inventory = PowerPointTemplate.Inspect(templatePath);
                PowerPointTemplateAssetInfo asset = Assert.Single(
                    inventory.Assets,
                    candidate => candidate.Name == "Extreme Logo");

                Assert.Null(asset.Bounds);
            } finally {
                Delete(templatePath);
            }
        }

        [Fact]
        public void CreateFromTemplate_RemovesSourceSlidesAndPreservesNamedLayout() {
            string templatePath = CreateTemplatePresentation();
            string outputPath = TempPath(".pptx");
            try {
                PowerPointTemplateInventory inventory = PowerPointTemplate.Inspect(templatePath);
                PowerPointTemplateLayoutInfo layout = inventory.Masters[0].Layouts[0];
                PowerPointTemplatePlaceholderInfo body = layout.ResolvePlaceholder("Executive Summary Body");

                using (PowerPointPresentation presentation = PowerPointTemplate.CreatePresentation(templatePath,
                           outputPath)) {
                    Assert.Empty(presentation.Slides);
                    PowerPointSlide slide = presentation.AddSlide(layout);
                    slide.AddTextToPlaceholder(body, "Generated from the corporate layout");
                    presentation.Save();
                    Assert.Empty(presentation.ValidateDocument());
                }

                using PowerPointPresentation reopened = PowerPointPresentation.Load(outputPath);
                Assert.Single(reopened.Slides);
                Assert.Equal("Corporate Test Theme", reopened.ThemeName);
                Assert.NotEmpty(reopened.GetSlideLayouts());
                Assert.Contains(reopened.Slides[0].TextBoxes,
                    textBox => textBox.Text == "Generated from the corporate layout" &&
                               textBox.Name == "Executive Summary Body");
            } finally {
                Delete(templatePath, outputPath);
            }
        }

        [Fact]
        public void CreateFromTemplate_SelectedRetentionIsDeterministicAndCanHideReferenceSlides() {
            string templatePath = CreateTemplatePresentation();
            string outputPath = TempPath(".pptx");
            try {
                var options = new PowerPointTemplateCreationOptions {
                    SlideRetention = PowerPointTemplateSlideRetention.Selected,
                    HideRetainedSourceSlides = true
                };
                options.SourceSlideIndexes.Add(1);

                using (PowerPointPresentation presentation = PowerPointTemplate.CreatePresentation(templatePath,
                           outputPath, options)) {
                    PowerPointSlide retained = Assert.Single(presentation.Slides);
                    Assert.True(retained.Hidden);
                    Assert.Contains(retained.TextBoxes, box => box.Text == "Source slide 2");
                    presentation.Save();
                }

                using PowerPointPresentation reopened = PowerPointPresentation.Load(outputPath);
                Assert.Single(reopened.Slides);
                Assert.True(reopened.Slides[0].Hidden);
            } finally {
                Delete(templatePath, outputPath);
            }
        }

        [Fact]
        public void CreateFromTemplate_DoesNotExposeCopiedTemplateWhenRetentionValidationFails() {
            string templatePath = CreateTemplatePresentation();
            string outputPath = TempPath(".pptx");
            try {
                var options = new PowerPointTemplateCreationOptions {
                    SlideRetention = PowerPointTemplateSlideRetention.Selected
                };
                options.SourceSlideIndexes.Add(99);

                Assert.Throws<ArgumentOutOfRangeException>(() =>
                    PowerPointTemplate.CreatePresentation(templatePath, outputPath, options));

                Assert.False(File.Exists(outputPath));
                string directory = Path.GetDirectoryName(outputPath)!;
                string prefix = "." + Path.GetFileNameWithoutExtension(outputPath) + ".";
                Assert.DoesNotContain(Directory.EnumerateFiles(directory), path =>
                    Path.GetFileName(path).StartsWith(prefix, StringComparison.Ordinal));
            } finally {
                Delete(templatePath, outputPath);
            }
        }

        [Fact]
        public void CreateFromTemplate_ConvertsPotxToEditablePptx() {
            string sourcePptx = CreateTemplatePresentation();
            string sourcePotx = TempPath(".potx");
            string outputPath = TempPath(".pptx");
            try {
                File.Copy(sourcePptx, sourcePotx);
                using (PresentationDocument template = PresentationDocument.Open(sourcePotx, true)) {
                    template.ChangeDocumentType(PresentationDocumentType.Template);
                    template.Save();
                }

                using (PowerPointPresentation presentation = PowerPointTemplate.CreatePresentation(sourcePotx,
                           outputPath)) {
                    Assert.Empty(presentation.Slides);
                    presentation.AddSlide();
                    presentation.Save();
                    Assert.Empty(presentation.ValidateDocument());
                }

                using PresentationDocument package = PresentationDocument.Open(outputPath, false);
                Assert.Equal(PresentationDocumentType.Presentation, package.DocumentType);
            } finally {
                Delete(sourcePptx, sourcePotx, outputPath);
            }
        }

        [Fact]
        public void TemplateDesigner_RendersSemanticPlanIntoMappedNamedLayout() {
            string templatePath = CreateTemplatePresentation();
            string outputPath = TempPath(".pptx");
            try {
                PowerPointTemplateInventory inventory = PowerPointTemplate.Inspect(templatePath);
                PowerPointTemplateLayoutInfo namedLayout = inventory.Masters[0].Layouts[0];
                var map = new PowerPointTemplateLayoutMap()
                    .Map(PowerPointDeckPlanSlideKind.Process, namedLayout);
                var plan = new PowerPointDeckPlan().AddProcess("Mapped process", null, new[] {
                    new PowerPointProcessStep("Discover", "Collect constraints."),
                    new PowerPointProcessStep("Deliver", "Ship with evidence.")
                });

                using (PowerPointPresentation presentation = PowerPointTemplate.CreatePresentation(templatePath,
                           outputPath)) {
                    PowerPointDeckComposer deck = presentation.UseTemplateDesigner(inventory, map,
                        "template-render", "delivery plan");
                    PowerPointSlide rendered = Assert.Single(deck.AddSlides(plan));
                    Assert.Equal(namedLayout.LayoutIndex, rendered.LayoutIndex);
                    Assert.Equal("Corporate Test Theme", presentation.ThemeName);
                    presentation.Save();
                    Assert.Empty(presentation.ValidateDocument());
                }

                using PowerPointPresentation reopened = PowerPointPresentation.Load(outputPath);
                PowerPointSlide slide = Assert.Single(reopened.Slides);
                Assert.Equal(namedLayout.LayoutIndex, slide.LayoutIndex);
                Assert.Contains(slide.TextBoxes, box => box.Text == "Mapped process");
            } finally {
                Delete(templatePath, outputPath);
            }
        }

        [Fact]
        public void ExistingPptxFixture_PreservesMastersLayoutsAndSlidesThroughTemplateCopy() {
            string fixturePath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
                "..", "..", "..", "..", "Assets", "PowerPointTemplates", "PowerPointWithTitle.pptx"));
            Assert.True(File.Exists(fixturePath), "Expected PowerPoint template fixture at " + fixturePath);
            string outputPath = TempPath(".pptx");
            try {
                PowerPointTemplateInventory source = PowerPointTemplate.Inspect(fixturePath);
                var options = new PowerPointTemplateCreationOptions {
                    SlideRetention = PowerPointTemplateSlideRetention.All
                };
                using (PowerPointPresentation copied = PowerPointTemplate.CreatePresentation(fixturePath,
                           outputPath, options)) {
                    PowerPointTemplateInventory result = PowerPointTemplate.Inspect(copied);
                    Assert.Equal(source.SourceSlideCount, result.SourceSlideCount);
                    Assert.Equal(source.Masters.Count, result.Masters.Count);
                    Assert.Equal(source.Masters.Sum(master => master.Layouts.Count),
                        result.Masters.Sum(master => master.Layouts.Count));
                    copied.Save();
                    Assert.Empty(copied.ValidateDocument());
                }
            } finally {
                Delete(outputPath);
            }
        }

        private static string CreateTemplatePresentation() {
            string path = TempPath(".pptx");
            using PowerPointPresentation presentation = PowerPointPresentation.Create(path);
            presentation.ThemeName = "Corporate Test Theme";
            presentation.SetThemeColor(PowerPointThemeColor.Accent1, "0B7FAB");
            presentation.SetThemeColor(PowerPointThemeColor.Accent2, "7C3AED");
            presentation.SetThemeColor(PowerPointThemeColor.Light2, "F5F7FA");
            presentation.SetThemeLatinFonts("Aptos Display", "Aptos");

            PowerPointTextBox primary = presentation.EnsureLayoutPlaceholderTextBox(0, 0,
                PlaceholderValues.Body, 20, PowerPointLayoutBox.FromCentimeters(1.5, 3.0, 18, 7),
                "Executive Summary Body");
            primary.Name = "Executive Summary Body";
            PowerPointTextBox secondary = presentation.EnsureLayoutPlaceholderTextBox(0, 0,
                PlaceholderValues.Body, 21, PowerPointLayoutBox.FromCentimeters(20, 3.0, 5, 7),
                "Supporting Body");
            secondary.Name = "Supporting Body";
            presentation.EnsureLayoutFooterPlaceholderTextBox(0, 0, "CONFIDENTIAL");

            for (int index = 1; index <= 3; index++) {
                PowerPointSlide slide = presentation.AddSlide();
                slide.AddTextBoxCm("Source slide " + index, 2, 2, 10, 1);
            }
            presentation.Save();
            return path;
        }

        private static string TempPath(string extension) =>
            Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + extension);

        private static void Delete(params string[] paths) {
            foreach (string path in paths) if (File.Exists(path)) File.Delete(path);
        }
    }
}
