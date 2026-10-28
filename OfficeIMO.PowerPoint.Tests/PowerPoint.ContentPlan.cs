using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class PowerPointContentPlanTests {
        private static PowerPointCompositionOptions Options() => PowerPointCompositionOptions.FromDesign(
            PowerPointDeckDesign.FromBrand("#008C95", "content-plan", headingFontName: "Aptos", bodyFontName: "Aptos"));

        [Fact]
        public void ImageCardsRetainThePictureAndBodyText() {
            string path = Path.Combine(Path.GetTempPath(), "officeimo-content-card-" + Guid.NewGuid().ToString("N") + ".png");
            var pixels = new byte[40 * 24 * 4];
            for (int index = 0; index < pixels.Length; index += 4) { pixels[index + 1] = 120; pixels[index + 2] = 180; pixels[index + 3] = 255; }
            File.WriteAllBytes(path, OfficePngWriter.EncodeRgba(40, 24, pixels));
            try {
                using var document = PowerPointPresentation.Create();
                var result = document.Compose(new PowerPointDeckPlan().AddContent("Evidence", new PowerPointLayoutBlock[] {
                    new PowerPointLayoutCard("Review", "Keep the image and the narrative.", image: path)
                }), Options());
                Assert.Single(result.Slides.SelectMany(slide => slide.Pictures));
                Assert.Contains(result.Slides.SelectMany(slide => slide.TextBoxes), box => box.Text.Contains("Keep the image"));
            } finally { File.Delete(path); }
        }

        [Fact]
        public void TwoContentKeepsItsLedeAndOptionalColumnTitles() {
            using var document = PowerPointPresentation.Create();
            var result = document.Compose(new PowerPointDeckPlan().AddContent("Evidence", new PowerPointLayoutBlock[] {
                new PowerPointLayoutText("Both sources matter."),
                new PowerPointLayoutColumn("Decisions", new PowerPointLayoutBlock[] { new PowerPointLayoutList(new[] { "Review first" }) }),
                new PowerPointLayoutColumn(null, new PowerPointLayoutBlock[] { new PowerPointLayoutList(new[] { "Then publish" }) })
            }), Options());
            var text = string.Join("\n", result.Slides.SelectMany(slide => slide.TextBoxes).Select(box => box.Text));
            Assert.Contains("Both sources matter.", text);
            Assert.Contains("Decisions", text);
            Assert.Contains("Then publish", text);
        }

        [Theory]
        [InlineData(PowerPointLayout.Auto)]
        [InlineData(PowerPointLayout.Closing)]
        public void ClosingKeepsBothSubtitleAndLede(PowerPointLayout layout) {
            using var document = PowerPointPresentation.Create();
            var plan = new PowerPointDeckPlan();
            for (int i = 0; i < 4; i++) plan.AddContent("Evidence " + i, Array.Empty<PowerPointLayoutBlock>());
            plan.AddContent("Actions", new PowerPointLayoutBlock[] {
                new PowerPointLayoutText("Keep this explanation."), new PowerPointLayoutList(new[] { "Publish the evidence" })
            }, layout, "Keep this subtitle.");
            var result = document.Compose(plan, Options());
            string text = string.Join("\n", result.Slides.Last().TextBoxes.Select(box => box.Text));
            Assert.Contains("Keep this explanation.", text);
            Assert.Contains("Keep this subtitle.", text);
            Assert.Contains("Publish the evidence", text);
        }

        [Fact]
        public void PreviewAndRenderUseTheLogicalPlanPositionAfterTableContinuations() {
            using var document = PowerPointPresentation.Create();
            var table = new PowerPointLayoutTable(new[] { "Evidence" },
                Enumerable.Range(1, 5).Select(index => (IReadOnlyList<string>)new[] { "Row " + index }).ToArray()) { MaximumRows = 1 };
            var plan = new PowerPointDeckPlan()
                .AddContent("Evidence", new PowerPointLayoutBlock[] { table })
                .AddContent("Actions", new PowerPointLayoutBlock[] { new PowerPointLayoutList(new[] { "Publish the evidence" }) });
            Assert.Equal("TitleAndContent", document.PreviewComposition(plan, Options()).Last().LayoutVariant);
            var result = document.Compose(plan, Options());
            Assert.Equal(6, result.Slides.Count);
            var body = Assert.Single(result.Slides.Last().TextBoxes, box => box.Text.Contains("Publish the evidence"));
            Assert.Contains(body.Paragraphs, paragraph => paragraph.BulletCharacter != null);
        }

        [Fact]
        public void UnsupportedSemanticMixturesFailBeforeAddingSlides() {
            using var document = PowerPointPresentation.Create();
            Assert.Throws<ArgumentException>(() => new PowerPointContentPlanSlide("Mixed", new PowerPointLayoutBlock[] {
                new PowerPointLayoutKpi("12", "Deliveries"), new PowerPointLayoutText("Keep the explanation")
            }));
            Assert.Throws<ArgumentException>(() => new PowerPointContentPlanSlide("Mixed", new PowerPointLayoutBlock[] {
                new PowerPointLayoutQuote("Keep the quote"), new PowerPointLayoutStep("Also keep the step")
            }));
            Assert.Empty(document.Slides);
        }

        [Fact]
        public void ClosingRequestsKeepNestedAndLongListsThroughTheAutomaticFallback() {
            using var document = PowerPointPresentation.Create();
            var items = Enumerable.Range(1, 6).Select(index => new PowerPointLayoutListItem("Action " + index, index == 2 ? 2 : 1));
            var result = document.Compose(new PowerPointDeckPlan().AddContent("Actions", new PowerPointLayoutBlock[] {
                new PowerPointLayoutList(items)
            }, PowerPointLayout.Closing), Options());
            string text = string.Join("\n", result.Slides.SelectMany(slide => slide.TextBoxes).Select(box => box.Text));
            Assert.Contains("Action 2", text);
            Assert.Contains("Action 6", text);
            Assert.Contains(result.Preflight.Findings, finding => finding.Code == "Layout.layout-content-mismatch");
        }

        [Theory]
        [InlineData(0, 1)]
        [InlineData(1, 2)]
        [InlineData(2, 1)]
        public void ContentTablesRejectMissingHeadersAndMismatchedRows(int headers, int cells) {
            Assert.Throws<ArgumentException>(() => new PowerPointLayoutTable(
                Enumerable.Range(0, headers).Select(index => "Column " + index).ToArray(),
                new IReadOnlyList<string>[] { Enumerable.Range(0, cells).Select(index => "Value " + index).ToArray() }));
        }

        [Theory]
        [InlineData(PowerPointLayout.Title)]
        [InlineData(PowerPointLayout.TitleAndContent)]
        public void IncompatibleExplicitLayoutsFallBackWithoutLosingTables(PowerPointLayout layout) {
            using var document = PowerPointPresentation.Create();
            var plan = new PowerPointDeckPlan().AddContent("Two tables", new PowerPointLayoutBlock[] {
                new PowerPointLayoutTable(new[] { "First" }, new IReadOnlyList<string>[] { new[] { "Alpha" } }),
                new PowerPointLayoutTable(new[] { "Second" }, new IReadOnlyList<string>[] { new[] { "Beta" } })
            }, layout);
            Assert.Equal("TwoContent", Assert.Single(document.PreviewComposition(plan, Options())).LayoutVariant);
            var result = document.Compose(plan, Options());
            Assert.Equal(2, result.Slides.Sum(slide => slide.Tables.Count()));
            Assert.Contains(result.Preflight.Findings, finding => finding.Code == "Layout.layout-content-mismatch");
        }

        [Fact]
        public void ColumnsRejectUnsupportedStacksInsteadOfDiscardingBlocks() {
            Assert.Throws<ArgumentException>(() => new PowerPointLayoutColumn("Mixed", new PowerPointLayoutBlock[] {
                new PowerPointLayoutList(new[] { "Keep this" }),
                new PowerPointLayoutTable(new[] { "Also keep" }, new IReadOnlyList<string>[] { new[] { "Evidence" } })
            }));
            Assert.Throws<ArgumentException>(() => new PowerPointLayoutColumn("Metric", new PowerPointLayoutBlock[] {
                new PowerPointLayoutKpi("12", "Deliveries")
            }));
        }

        [Fact]
        public void QuoteRequestsKeepMultipleParagraphsThroughTheAutomaticFallback() {
            using var document = PowerPointPresentation.Create();
            var result = document.Compose(new PowerPointDeckPlan().AddContent(null, new PowerPointLayoutBlock[] {
                new PowerPointLayoutText(new[] { "First observation.", "Second observation." })
            }, PowerPointLayout.Quote), Options());
            Assert.Contains(result.Slides.SelectMany(slide => slide.TextBoxes), box => box.Text.Contains("Second observation."));
        }

        [Fact]
        public void ContentPreviewsSelectTheSameKpiLayoutWithoutAddingSlides() {
            using var document = PowerPointPresentation.Create();
            var plan = new PowerPointDeckPlan().AddContent("Results", new PowerPointLayoutBlock[] { new PowerPointLayoutKpi("12", "Deliveries") });
            Assert.Equal("Kpi", Assert.Single(document.PreviewComposition(plan, Options())).LayoutVariant);
            Assert.Empty(document.Slides);
        }

        [Theory]
        [InlineData(OfficeColorRole.Accent4, "accent4")]
        [InlineData(OfficeColorRole.Accent5, "accent5")]
        [InlineData(OfficeColorRole.Accent6, "accent6")]
        [InlineData(OfficeColorRole.TextMuted, "dk2")]
        public void ContentColoursRetainTheirExactNativeThemeSlot(OfficeColorRole role, string expected) {
            using var document = PowerPointPresentation.Create();
            var plan = new PowerPointDeckPlan().AddContent("Results", new PowerPointLayoutBlock[] {
                new PowerPointLayoutKpi("12", "Deliveries", "+2", OfficeColor.Parse(role.ToString()))
            });
            document.Compose(plan, Options());
            var fill = SavedMetricChangeFill(document);
            Assert.Equal(expected, Assert.Single(fill.Elements<DocumentFormat.OpenXml.Drawing.SchemeColor>()).Val!.InnerText);
        }

        [Fact]
        public void ExplicitContentColourStaysLiteralWhenItEqualsAThemeColour() {
            using var document = PowerPointPresentation.Create();
            var options = Options();
            document.Compose(new PowerPointDeckPlan(), options);
            string hex = document.GetThemeColor(PowerPointThemeColor.Accent1)!;
            document.Compose(new PowerPointDeckPlan().AddContent("Results", new PowerPointLayoutBlock[] {
                new PowerPointLayoutKpi("12", "Deliveries", "+2", OfficeColor.Parse(hex))
            }), options);
            var fill = SavedMetricChangeFill(document);
            Assert.Equal(hex, Assert.Single(fill.Elements<DocumentFormat.OpenXml.Drawing.RgbColorModelHex>()).Val!.Value);
            Assert.Empty(fill.Elements<DocumentFormat.OpenXml.Drawing.SchemeColor>());
        }

        private static DocumentFormat.OpenXml.Drawing.SolidFill SavedMetricChangeFill(PowerPointPresentation document) {
            using var stream = new MemoryStream();
            document.Save(stream);
            stream.Position = 0;
            using var package = DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(stream, false);
            var shape = Assert.Single(package.PresentationPart!.SlideParts.SelectMany(part => part.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Shape>()),
                item => item.InnerText.Contains("+2"));
            return (DocumentFormat.OpenXml.Drawing.SolidFill)Assert.Single(shape.Descendants<DocumentFormat.OpenXml.Drawing.SolidFill>()).CloneNode(true);
        }

        [Fact]
        public void ContentColoursKeepThemeRolesAndNativeAlpha() {
            using var document = PowerPointPresentation.Create();
            document.Compose(new PowerPointDeckPlan()
                .AddContent("Results", new PowerPointLayoutBlock[] { new PowerPointLayoutKpi("12", "Deliveries", "+2", OfficeColor.Accent2) })
                .AddContent("Review", new PowerPointLayoutBlock[] {
                    new PowerPointLayoutCard("Evidence", "Ready for review", color: OfficeColor.Parse("#12345680"))
                }), Options());
            using var stream = new MemoryStream();
            document.Save(stream);
            stream.Position = 0;
            using var package = DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(stream, false);
            var slides = package.PresentationPart!.SlideParts.ToArray();
            Assert.Contains(slides.SelectMany(slide => slide.Slide.Descendants<DocumentFormat.OpenXml.Drawing.SchemeColor>()),
                color => color.Val!.Value == DocumentFormat.OpenXml.Drawing.SchemeColorValues.Accent2);
            var literal = Assert.Single(slides.SelectMany(slide => slide.Slide.Descendants<DocumentFormat.OpenXml.Drawing.RgbColorModelHex>()),
                color => color.Val!.Value == "123456");
            Assert.Equal(50196, literal.GetFirstChild<DocumentFormat.OpenXml.Drawing.Alpha>()!.Val!.Value);
        }

        [Fact]
        public void PreflightAndTableSnapshotsResolveTheNativeThemeReferences() {
            using var document = PowerPointPresentation.Create();
            var result = document.Compose(new PowerPointDeckPlan()
                .AddContent("Quarterly delivery review", Array.Empty<PowerPointLayoutBlock>(), subtitle: "Evidence, results and next steps")
                .AddContent("Evidence", new PowerPointLayoutBlock[] {
                    new PowerPointLayoutTable(new[] { "Item", "Score" }, new IReadOnlyList<string>[] { new[] { "Delivery", "96" } })
                }), Options());
            Assert.DoesNotContain(result.Preflight.Findings, finding => finding.Code == "text-clipped");
            var snapshot = result.Slides[1].CreateVisualSnapshot();
            Assert.DoesNotContain(snapshot.Diagnostics, diagnostic => diagnostic.Message.Contains("+mn-lt"));
            Assert.DoesNotContain(snapshot.TextSpans, span => span.BaseFont != null && span.BaseFont.StartsWith("+"));
            Assert.Contains(snapshot.TextSpans, span => span.Text.Contains("Delivery"));
        }

        [Fact]
        public void UntitledContentCanSelectAQuoteWithoutAnArtificialVisibleTitle() {
            using var document = PowerPointPresentation.Create();
            var result = document.Compose(new PowerPointDeckPlan().AddContent(null,
                new PowerPointLayoutBlock[] { new PowerPointLayoutText("Evidence before decisions.") }), Options());
            var slide = Assert.Single(result.Slides);
            Assert.Contains(slide.TextBoxes, box => box.Text.Contains("Evidence before decisions."));
            Assert.Contains(slide.TextBoxes, box => box.Text == "\u201c");
        }

        [Fact]
        public void PublicContentComposesSemanticLayoutsAndKeepsSpeakerNotesSeparate() {
            var quoted = new PowerPointContentPlanSlide("A useful observation", new PowerPointLayoutBlock[] {
                new PowerPointLayoutQuote("Clear evidence makes decisions easier.", "Review team")
            }) { Notes = "Discuss the evidence before the decision." };
            var plan = new PowerPointDeckPlan()
                .AddContent("Quarterly review", Array.Empty<PowerPointLayoutBlock>())
                .AddContent("Results", new PowerPointLayoutBlock[] {
                    new PowerPointLayoutKpi("96%", "On time", "+4 points"),
                    new PowerPointLayoutKpi("42", "Deliveries")
                })
                .Add(quoted)
                .AddContent("Delivery", new PowerPointLayoutBlock[] {
                    new PowerPointLayoutStep("Discover", "Check the evidence"),
                    new PowerPointLayoutStep("Deliver", "Verify the result")
                })
                .AddContent("Next steps", new PowerPointLayoutBlock[] { new PowerPointLayoutList(new[] { "Share the findings" }) });
            using var document = PowerPointPresentation.Create();
            var result = document.Compose(plan, Options());
            Assert.Equal(5, result.Slides.Count);
            Assert.Equal(document.Slides.ToArray(), result.Slides.ToArray());
            Assert.Contains(result.Slides[1].TextBoxes, box => box.Text == "96%");
            Assert.Contains(result.Slides[2].TextBoxes, box => box.Text.Contains("Clear evidence"));
            Assert.Equal(quoted.Notes, result.Slides[2].Notes.Text);
            Assert.DoesNotContain(result.Slides[2].TextBoxes, box => box.Text.Contains("Discuss the evidence"));
            Assert.Contains(result.Slides[3].TextBoxes, box => box.Text.Contains("Discover"));
            // Auto uses the complete plan's last position for its closing-slide choice.
            Assert.Contains(result.Slides[4].TextBoxes, box => box.Text == "Next steps");
            using var stream = new MemoryStream();
            document.Save(stream);
            stream.Position = 0;
            using var loaded = PowerPointPresentation.Load(stream);
            Assert.Equal(5, loaded.Slides.Count);
            Assert.Equal(quoted.Notes, loaded.Slides[2].Notes.Text);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void CompositionReturnsEveryPhysicalTableContinuation(bool contentPlan) {
            var headers = new[] { "Item", "Value" };
            var rows = Enumerable.Range(1, 35)
                .Select(index => (IReadOnlyList<string>)new[] { "Item " + index, index.ToString() }).ToArray();
            var plan = contentPlan
                ? new PowerPointDeckPlan().AddContent("Evidence", new PowerPointLayoutBlock[] { new PowerPointLayoutTable(headers, rows) })
                : new PowerPointDeckPlan().AddAppendixTable("Evidence", null, new PowerPointTableData(headers, rows));
            using var document = PowerPointPresentation.Create();
            var result = document.Compose(plan, Options());
            Assert.True(result.Slides.Count > 1);
            Assert.Equal(document.Slides.ToArray(), result.Slides.ToArray());
            Assert.Equal(35, result.Slides.Sum(slide => slide.Tables.Single().Rows - 1));
            Assert.All(result.Slides, slide => Assert.Equal("Item", slide.Tables.Single().GetCell(0, 0).Text));
            using var stream = new MemoryStream();
            document.Save(stream);
            stream.Position = 0;
            using var loaded = PowerPointPresentation.Load(stream);
            Assert.Equal(result.Slides.Count, loaded.Slides.Count);
            Assert.Equal(35, loaded.Slides.Sum(slide => slide.Tables.Single().Rows - 1));
        }
    }
}
