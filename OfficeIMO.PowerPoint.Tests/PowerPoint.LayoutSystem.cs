using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.Tests {
    public class PowerPointLayoutSystem {
        private static string TempImage(int width, int height) {
            string path = Path.Combine(Path.GetTempPath(), "officeimo-layout-" + Guid.NewGuid().ToString("N") + ".png");
            var rgba = new byte[width * height * 4];
            for (int i = 0; i < rgba.Length; i += 4) {
                rgba[i] = 40;
                rgba[i + 1] = 90;
                rgba[i + 2] = 200;
                rgba[i + 3] = 255;
            }

            File.WriteAllBytes(path, OfficePngWriter.EncodeRgba(width, height, rgba));
            return path;
        }

        private static IEnumerable<P.Shape> Shapes(PowerPointSlide slide) =>
            slide.SlidePart.Slide!.CommonSlideData!.ShapeTree!.Elements<P.Shape>();

        private static P.PlaceholderShape? Placeholder(P.Shape shape) =>
            shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape;

        private static double Points(long emus) => emus / 12700D;

        private static PowerPointLayoutTable Table(int rows, int columns = 4) =>
            new(Enumerable.Range(1, columns).Select(c => "Column " + c).ToList(),
                Enumerable.Range(1, rows).Select(r => (IReadOnlyList<string>)Enumerable.Range(1, columns)
                    .Select(c => c == columns ? (r * 10).ToString() : "Row " + r + " value " + c).ToList()).ToList());

        [Fact]
        public void NewPresentation_WritesThemeTextStylesAndPlaceholders() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            PresentationPart part = presentation.OpenXmlDocument.PresentationPart!;
            SlideMasterPart master = part.SlideMasterParts.First();
            P.TextStyles styles = master.SlideMaster!.TextStyles!;
            PowerPointLayoutTheme theme = PowerPointLayoutTheme.From(OfficeTheme.Default);

            A.DefaultRunProperties title = styles.TitleStyle!.GetFirstChild<A.Level1ParagraphProperties>()!.GetFirstChild<A.DefaultRunProperties>()!;
            Assert.Equal((int)(theme[PowerPointTextRole.Title].Size * 100), title.FontSize!.Value);
            Assert.Equal("+mj-lt", title.GetFirstChild<A.LatinFont>()!.Typeface!.Value);
            A.Level1ParagraphProperties body = styles.BodyStyle!.GetFirstChild<A.Level1ParagraphProperties>()!;
            Assert.Equal((int)(theme[PowerPointTextRole.Body].Size * 100), body.GetFirstChild<A.DefaultRunProperties>()!.FontSize!.Value);
            Assert.NotNull(body.GetFirstChild<A.CharacterBullet>());
            Assert.NotNull(part.Presentation!.DefaultTextStyle!.GetFirstChild<A.Level1ParagraphProperties>()!.GetFirstChild<A.DefaultRunProperties>()!.FontSize);

            var types = master.SlideMaster.CommonSlideData!.ShapeTree!.Elements<P.Shape>()
                .Select(s => Placeholder(s)?.Type?.Value).ToList();
            Assert.Contains(P.PlaceholderValues.Title, types);
            Assert.Contains(P.PlaceholderValues.Body, types);
            Assert.Contains(P.PlaceholderValues.Footer, types);
            Assert.Contains(P.PlaceholderValues.SlideNumber, types);

            // Layout placeholders use the 16:9 grid rather than 4:3 coordinates.
            SlideLayoutPart content = master.SlideLayoutParts.First(l => l.SlideLayout!.CommonSlideData!.Name == "Title and Content");
            P.Shape titleShape = content.SlideLayout!.CommonSlideData!.ShapeTree!.Elements<P.Shape>().First(s => Placeholder(s)?.Type?.Value == P.PlaceholderValues.Title);
            double width = Points(titleShape.ShapeProperties!.Transform2D!.Extents!.Cx!.Value);
            var grid = new PowerPointLayoutGrid(960, 540, theme.MarginLeft, theme.MarginRight, theme.MarginTop);
            Assert.Equal(grid.W(10), width, 1);
        }

        [Fact]
        public void ThemeFonts_CarryMatchingPanoseAndFamily() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            A.FontScheme scheme = presentation.OpenXmlDocument.PresentationPart!.SlideMasterParts.First().ThemePart!.Theme!.ThemeElements!.FontScheme!;
            Assert.Equal("Aptos", scheme.MajorFont!.LatinFont!.Typeface!.Value);
            Assert.Equal("020B0004020202020204", scheme.MajorFont.LatinFont.Panose!.Value);

            presentation.SetThemeLatinFonts("Georgia", "Segoe UI");
            Assert.Equal("02040502050405020303", scheme.MajorFont.LatinFont!.Panose!.Value);
            Assert.Equal("020B0502040204020203", scheme.MinorFont!.LatinFont!.Panose!.Value);

            presentation.SetThemeLatinFonts("Unknown Brand Sans", "Aptos");
            Assert.Null(scheme.MajorFont.LatinFont!.Panose);
            Assert.Equal(0x22, (byte)scheme.MajorFont.LatinFont.PitchFamily!.Value);
        }

        [Fact]
        public void CreateOptionsTheme_DrivesThemePartAndTextStyles() {
            OfficeTheme theme = OfficeTheme.Modern.WithFonts(new OfficeThemeFonts(
                new OfficeThemeFont("Georgia", new[] { "serif" }), new OfficeThemeFont("Arial", new[] { "sans-serif" })));
            using PowerPointPresentation presentation = PowerPointPresentation.Create(new PowerPointCreateOptions { Theme = theme });
            Assert.Equal("Georgia", presentation.GetThemeFonts().MajorLatin);
            Assert.Equal("Arial", presentation.GetThemeFonts().MinorLatin);
            Assert.Equal(theme.Palette.Accent1.ToRgbHex(), presentation.GetThemeColor(PowerPointThemeColor.Accent1));
        }

        [Fact]
        public void AddPicture_KeepsImageAspectRatioInsideTheBox() {
            string portrait = TempImage(70, 100);
            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create();
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointPicture picture = slide.AddPicturePoints(portrait, 100, 100, 300, 200);
                Assert.Equal(0.7, picture.Width / (double)picture.Height, 2);
                Assert.Equal(200, Points(picture.Height), 1);
                Assert.Equal(100 + (300 - 140) / 2D, Points(picture.Left), 1);

                // Cropping to fill crops the long axis of a portrait image: top and bottom.
                picture.Width = PowerPointUnits.FromPoints(300);
                picture.Height = PowerPointUnits.FromPoints(200);
                picture.FitToBox(crop: true);
                Assert.True(picture.CropTopRatio > 0.2);
                Assert.Equal(picture.CropTopRatio, picture.CropBottomRatio, 3);
                Assert.Equal(0, picture.CropLeftRatio, 3);
            } finally {
                File.Delete(portrait);
            }
        }

        [Fact]
        public void TableFrame_FollowsGridWhenColumnsAndRowsChange() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointTable table = slide.AddTablePoints(3, 3, 40, 40, 875, 390);
            table.SetColumnWidthsPoints(200, 180, 180);
            for (int r = 0; r < table.Rows; r++) table.SetRowHeightPoints(r, 30);
            Assert.Equal(560, Points(table.Width), 1);
            Assert.Equal(90, Points(table.Height), 1);
            table.AddRow();
            Assert.Equal(Enumerable.Range(0, table.Rows).Sum(r => table.GetRowHeight(r)), table.Height);
        }

        [Fact]
        public void LongTable_SplitsIntoBalancedSlidesWithRepeatedHeaderAndRange() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            var slide = new PowerPointLayoutSlide { Title = "Workstreams" };
            slide.Blocks.Add(Table(17));
            IReadOnlyList<PowerPointSlide> slides = PowerPointLayoutEngine.AddSlides(presentation, slide);

            Assert.InRange(slides.Count, 2, 3);
            List<PowerPointTable> tables = slides.Select(s => s.Tables.Single()).ToList();
            int[] dataRows = tables.Select(t => t.Rows - 1).ToArray();
            Assert.Equal(17, dataRows.Sum());
            Assert.True(dataRows.Max() - dataRows.Min() <= 1, "Parts are not balanced: " + string.Join(" + ", dataRows));
            Assert.All(tables, t => Assert.Equal("Column 1", t.GetCell(0, 0).Text)); // header repeated
            int first = 1;
            for (int i = 0; i < slides.Count; i++) {
                int last = first + dataRows[i] - 1;
                Assert.Contains(slides[i].TextBoxes, t => t.Text == "Rows " + first + "\u2013" + last + " of 17");
                first = last + 1;
            }
            Assert.Contains(slides[1].TextBoxes, t => t.Text.Contains("(continued"));

            var grid = PowerPointLayoutGrid.For(presentation, presentation.LayoutTheme);
            foreach (PowerPointTable table in tables) {
                Assert.True(Points(table.Top + table.Height) <= grid.LiveBottom + 0.5, "The table runs into the footer.");
                // Continuation rows keep the table's own row height.
                Assert.Equal(tables[0].GetRowHeight(1), table.GetRowHeight(1));
                A.RunProperties run = table.TableElement.Descendants<A.RunProperties>().Skip(table.Columns).First();
                Assert.Equal((int)(presentation.LayoutTheme[PowerPointTextRole.TableBody].Size * 100), run.FontSize!.Value);
            }
        }

        [Fact]
        public void LongList_ShrinksThenSplitsEvenlyAndContinuesNumbering() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            var slide = new PowerPointLayoutSlide { Title = "Habits that keep generated decks readable" };
            slide.Blocks.Add(new PowerPointLayoutList(Enumerable.Range(1, 14).Select(i => "Keep one idea on each slide, item " + i), numbered: true));
            IReadOnlyList<PowerPointSlide> slides = PowerPointLayoutEngine.AddSlides(presentation, slide);

            Assert.Equal(2, slides.Count);
            List<P.Shape> lists = slides.Select(s => Shapes(s).Single(x => x.NonVisualShapeProperties!.NonVisualDrawingProperties!.Name!.Value!.StartsWith("List"))).ToList();
            Assert.Equal(7, lists[0].Descendants<A.Paragraph>().Count());
            Assert.Equal(7, lists[1].Descendants<A.Paragraph>().Count());
            Assert.Equal(8, lists[1].Descendants<A.AutoNumberedBullet>().First().StartAt!.Value);
            // Both parts use the same size, never below the minimum.
            int[] sizes = lists.Select(l => l.Descendants<A.RunProperties>().First().FontSize!.Value).ToArray();
            Assert.Equal(sizes[0], sizes[1]);
            Assert.True(sizes[0] >= 1600);
        }

        [Fact]
        public void LongTitle_SplitsAtColonIntoTitleAndSubtitle() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            var slide = new PowerPointLayoutSlide {
                Title = "Remote staff intake pilot: why it should continue through the first quarter of next year before any wider rollout is approved"
            };
            slide.Blocks.Add(new PowerPointLayoutList(new[] { "Demand is settling", "Two of three measures are met" }));
            PowerPointSlide added = PowerPointLayoutEngine.AddSlides(presentation, slide).Single();
            P.Shape title = Shapes(added).Single(s => Placeholder(s)?.Type?.Value == P.PlaceholderValues.Title);
            Assert.Equal("Remote staff intake pilot", title.InnerText);
            Assert.Contains(added.TextBoxes, t => t.Text.StartsWith("Why it should continue", StringComparison.Ordinal));
        }

        [Theory]
        [InlineData("kpi", "Kpi")]
        [InlineData("cards", "Cards")]
        [InlineData("logos", "Logos")]
        [InlineData("case", "CaseStudy")]
        [InlineData("steps", "Process")]
        [InlineData("locations", "Coverage")]
        [InlineData("columns", "Comparison")]
        [InlineData("quote", "Quote")]
        [InlineData("title-only", "Section")]
        [InlineData("one", "TitleAndContent")]
        [InlineData("two", "TwoContent")]
        [InlineData("lede", "TitleAndContent")]
        public void AutomaticSelection_FollowsTheContent(string content, string expectedLayout) {
            var expected = (PowerPointLayout)Enum.Parse(typeof(PowerPointLayout), expectedLayout);
            var slide = new PowerPointLayoutSlide { Title = "Slide" };
            var list = new PowerPointLayoutList(new[] { "One", "Two" });
            switch (content) {
                case "kpi": slide.Blocks.Add(new PowerPointLayoutKpi("1", "One")); break;
                case "cards": slide.Blocks.Add(new PowerPointLayoutCard("A", "Text")); break;
                case "logos": slide.Blocks.Add(new PowerPointLayoutCard("A", image: "logo.png")); break;
                case "case": slide.Blocks.Add(new PowerPointLayoutCard("A", "Text")); slide.Blocks.Add(new PowerPointLayoutKpi("1", "One")); break;
                case "steps": slide.Blocks.Add(new PowerPointLayoutStep("A")); break;
                case "locations": slide.Blocks.Add(new PowerPointLayoutLocation("Dublin", 0.4, 0.3)); break;
                case "columns":
                    slide.Blocks.Add(new PowerPointLayoutColumn("Before", new[] { list }));
                    slide.Blocks.Add(new PowerPointLayoutColumn("After", new[] { list }));
                    break;
                case "quote": slide.Blocks.Add(new PowerPointLayoutQuote("Quote", "Someone")); break;
                case "one": slide.Blocks.Add(list); break;
                case "two": slide.Blocks.Add(list); slide.Blocks.Add(new PowerPointLayoutText("Text")); break;
                case "lede": slide.Blocks.Add(new PowerPointLayoutText("A lede.")); slide.Blocks.Add(list); break;
            }

            Assert.Equal(expected, PowerPointLayoutEngine.Select(slide, isFirst: false, isLast: false, deckSize: 3));
            Assert.Equal(PowerPointLayout.Title, PowerPointLayoutEngine.Select(new PowerPointLayoutSlide { Title = "Deck" }, true, false, 1));
            Assert.Equal(PowerPointLayout.Closing, PowerPointLayoutEngine.Select(new PowerPointLayoutSlide { Title = "Thanks" }, false, true, 6));
        }

        [Fact]
        public void ContentSlide_UsesTitlePlaceholderSectionEyebrowAndNativeFooter() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            presentation.LayoutDeck.FooterText = "Northwind review";
            PowerPointLayoutEngine.AddSlides(presentation, new PowerPointLayoutSlide { Title = "Quarterly review" });
            PowerPointLayoutEngine.AddSlides(presentation, new PowerPointLayoutSlide { Title = "Performance" });
            var kpis = new PowerPointLayoutSlide { Title = "At a glance" };
            kpis.Blocks.Add(new PowerPointLayoutKpi("$5.7m", "Revenue", "+12%"));
            kpis.Blocks.Add(new PowerPointLayoutKpi("2.1%", "Churn", "-0.4 pt"));
            PowerPointSlide slide = PowerPointLayoutEngine.AddSlides(presentation, kpis).Single();

            List<P.Shape> shapes = Shapes(slide).ToList();
            Assert.Single(shapes, s => Placeholder(s)?.Type?.Value == P.PlaceholderValues.Title);
            P.Shape footer = shapes.Single(s => Placeholder(s)?.Type?.Value == P.PlaceholderValues.Footer);
            Assert.Equal("Northwind review", footer.InnerText);
            P.Shape number = shapes.Single(s => Placeholder(s)?.Type?.Value == P.PlaceholderValues.SlideNumber);
            Assert.Equal("slidenum", number.Descendants<A.Field>().Single().Type!.Value);
            Assert.Contains(slide.TextBoxes, t => t.Text == "PERFORMANCE");
            Assert.Contains(slide.TextBoxes, t => t.Text.StartsWith("▲", StringComparison.Ordinal));
            Assert.Contains(slide.TextBoxes, t => t.Text.StartsWith("▼", StringComparison.Ordinal));
            Assert.Equal("Title Slide", presentation.Slides[0].SlideLayoutName);
            Assert.Equal("Section Header", presentation.Slides[1].SlideLayoutName);
            Assert.DoesNotContain(Shapes(presentation.Slides[0]), s => Placeholder(s)?.Type?.Value == P.PlaceholderValues.Footer);
        }

        [Fact]
        public void TextFrames_AreNeverShorterThanTheirLines() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            var plan = new PowerPointDeckPlan()
                .AddSection("Operating model", "One semantic plan and one renderer")
                .AddCardGrid("Ownership", "Concrete objects remain editable", new[] {
                    new PowerPointCardContent("Lifecycle", new[] { "Create", "Load", "Save" }),
                    new PowerPointCardContent("Editing", new[] { "Slides", "Shapes", "Content" }),
                    new PowerPointCardContent("Review", new[] { "Preflight", "Accessibility" })
                })
                .AddProcess("Delivery", null, Enumerable.Range(1, 8).Select(i => new PowerPointProcessStep("Phase " + i, "Do the work for phase " + i + ".")));
            PowerPointCompositionResult result = presentation.Compose(plan,
                PowerPointCompositionOptions.FromDesign(PowerPointDeckDesign.FromBrand("#008C95", "layout-tests")));

            Assert.Equal(3, result.Slides.Count); // eight steps fit one Process slide
            foreach (PowerPointSlide slide in result.Slides) {
                foreach (PowerPointTextBox box in slide.TextBoxes.Where(t => !string.IsNullOrWhiteSpace(t.Text))) {
                    P.Shape shape = Shapes(slide).Single(s => s.NonVisualShapeProperties!.NonVisualDrawingProperties!.Name == box.Name);
                    int largest = shape.Descendants<A.RunProperties>().Select(r => r.FontSize?.Value ?? 0).DefaultIfEmpty(0).Max();
                    Assert.True(Points(box.Height) + 0.5 >= largest / 100D, box.Name + " is shorter than one line of its text.");
                    Assert.True(largest == 0 || largest >= 1000, box.Name + " uses text below 10 pt.");
                }
            }
        }

        [Fact]
        public void PlanContinuations_SplitBalanced() {
            var plan = new PowerPointDeckPlan().AddCardGrid("Capabilities", null,
                Enumerable.Range(1, 8).Select(i => new PowerPointCardContent("Card " + i, new[] { "Item" })));
            PowerPointDeckPlan expanded = plan.WithContinuations();
            Assert.Equal(new[] { 4, 4 }, expanded.Slides.Cast<PowerPointCardGridPlanSlide>().Select(s => s.Cards.Count).ToArray());
            Assert.EndsWith("(continued)", expanded.Slides[1].Title, StringComparison.Ordinal);
        }

        [Fact]
        public void LogoNames_ShrinkSoNoWordBreaks() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            var slide = new PowerPointLayoutSlide { Title = "Partners", Layout = PowerPointLayout.Logos };
            foreach (string name in new[] { "Brother", "Samsung", "Hewlett Packard Enterprise", "Lenovo", "Dell", "Cisco", "Fortinet", "Microsoft", "Amazon" }) {
                slide.Blocks.Add(new PowerPointLayoutCard(name, image: "missing.png"));
            }

            PowerPointSlide added = PowerPointLayoutEngine.AddSlides(presentation, slide).Single();
            foreach (P.Shape shape in Shapes(added).Where(s => s.NonVisualShapeProperties!.NonVisualDrawingProperties!.Name!.Value!.StartsWith("Logo Name"))) {
                double width = Points(shape.ShapeProperties!.Transform2D!.Extents!.Cx!.Value);
                double size = shape.Descendants<A.RunProperties>().First().FontSize!.Value / 100D;
                Assert.True(PowerPointTextMeasure.LongestWord(shape.InnerText, size, "Aptos", bold: true) <= width);
            }
        }
    
        [Fact]
        public void DesignerDeck_KeepsContentInsideTheLiveArea() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            var rows = Enumerable.Range(1, 20).Select(i => (IEnumerable<string>)new[] { "Workstream " + i, i % 3 == 0 ? "At risk" : "On track", (50 + i) + "%" });
            var plan = new PowerPointDeckPlan()
                .AddSection("Executive delivery review", "Editable, measured and ready for inspection")
                .AddExecutiveSummary("Executive summary", "A decision-ready opening", new PowerPointExecutiveSummaryContent(
                    new[] { new PowerPointMetric("14", "proof slides"), new PowerPointMetric("0", "hidden rows") },
                    new[] { new PowerPointCardContent("Decision", new[] { "Use semantic story families" }) }))
                .AddComparison("Implementation choice", "Two options", new[] {
                    new PowerPointComparisonItem("Shared core", "One owner", new[] { "Consistent", "Testable" }, new[] { "Needs contracts" }),
                    new PowerPointComparisonItem("Local helpers", "Per project", new[] { "Fast locally" }, new[] { "Drifts" })
                })
                .AddArchitecture("End-to-end ownership", null, new PowerPointArchitectureContent(new[] {
                    new PowerPointArchitectureNode("core", "Drawing core", "Shared semantics", "Core"),
                    new PowerPointArchitectureNode("ppt", "PowerPoint", "Native authoring", "Surfaces"),
                    new PowerPointArchitectureNode("md", "Markup", "Thin adapter", "Surfaces")
                }, new[] { new PowerPointArchitectureEdge("core", "ppt", "drives"), new PowerPointArchitectureEdge("core", "md", "shares") }))
                .AddProcess("Delivery path", null, Enumerable.Range(1, 8).Select(i => new PowerPointProcessStep("Phase " + i, "Evidence and owner for phase " + i + ".")))
                .AddCardGrid("Decision signals", null, Enumerable.Range(1, 8).Select(i => new PowerPointCardContent("Signal " + i, new[] { "Owner assigned", "Evidence available" })))
                .AddAppendixTable("Delivery detail", "Rows continue", new PowerPointTableData(new[] { "Workstream", "Status", "Complete" }, rows))
                .AddClosing("Next decision", new PowerPointClosingContent("Automation people can trust.", "Approve the next release."));
            PowerPointCompositionResult result = presentation.Compose(plan,
                PowerPointCompositionOptions.FromDesign(PowerPointDeckDesign.FromBrand("#085B7B", "live-area")));

            var grid = PowerPointLayoutGrid.For(presentation, presentation.LayoutTheme);
            foreach (PowerPointSlide slide in result.Slides) {
                foreach (P.Shape shape in Shapes(slide)) {
                    string name = shape.NonVisualShapeProperties!.NonVisualDrawingProperties!.Name!.Value!;
                    A.Transform2D? xfrm = shape.ShapeProperties?.Transform2D;
                    if (xfrm?.Offset == null || name.StartsWith("Section Number", StringComparison.Ordinal)) continue;
                    double bottom = Points(xfrm.Offset.Y!.Value + xfrm.Extents!.Cy!.Value);
                    double right = Points(xfrm.Offset.X!.Value + xfrm.Extents.Cx!.Value);
                    Assert.True(right <= grid.SlideWidth + 0.5 && bottom <= grid.SlideHeight + 0.5, name + " leaves the slide.");
                    P.PlaceholderShape? ph = Placeholder(shape);
                    bool footer = ph?.Type?.Value == P.PlaceholderValues.Footer || ph?.Type?.Value == P.PlaceholderValues.SlideNumber;
                    if (!footer && !name.StartsWith("Section", StringComparison.Ordinal)) {
                        Assert.True(bottom <= grid.LiveBottom + 0.5, name + " on '" + slide.SlideLayoutName + "' runs into the footer band (" + bottom + ").");
                    }
                }

                foreach (PowerPointTable table in slide.Tables) {
                    Assert.True(Points(table.Top + table.Height) <= grid.LiveBottom + 0.5, "A table runs into the footer band.");
                }
            }

            Assert.Empty(presentation.ValidateDocument());
        }
    
        [Fact]
        public void LayoutCharts_DropTheFrameAndKeepSliceLabelsReadable() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            var slide = new PowerPointLayoutSlide { Title = "Ticket mix by channel" };
            var data = new OfficeChartData(new[] { "Email", "Phone", "Chat", "Walk-up" },
                new[] { new OfficeChartSeries("Tickets", new double[] { 42, 28, 18, 12 }) });
            slide.Blocks.Add(new PowerPointLayoutChart(OfficeChartKind.Doughnut, data));
            PowerPointSlide added = PowerPointLayoutEngine.AddSlides(presentation, slide).Single();
            PowerPointChart chart = added.Charts.Single();
            DocumentFormat.OpenXml.Drawing.Charts.ChartSpace space = chart.GetChart().Parent as DocumentFormat.OpenXml.Drawing.Charts.ChartSpace
                ?? throw new InvalidOperationException("Chart space missing.");
            Assert.NotNull(space.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.ShapeProperties>()?.GetFirstChild<A.NoFill>());

            // Every slice carries its share. One label colour serves all slices (per-slice label formatting does
            // not survive every renderer), so it must reach large-text contrast (3:1 at 14 pt bold) on each slice.
            var dataLabels = chart.GetChart().Descendants<DocumentFormat.OpenXml.Drawing.Charts.DataLabels>().ToList();
            Assert.Contains(dataLabels, l => l.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.ShowPercent>()?.Val?.Value == true);
            Assert.Empty(chart.GetChart().Descendants<DocumentFormat.OpenXml.Drawing.Charts.DataLabel>());
            var run = dataLabels.SelectMany(l => l.Descendants<A.DefaultRunProperties>()).First();
            Assert.Equal(1400, run.FontSize!.Value);
            string color = run.Descendants<A.RgbColorModelHex>().Last().Val!.Value!;
            for (int index = 0; index < 4; index++) {
                Assert.True(PowerPointLayoutTheme.Contrast(color, presentation.LayoutTheme.Accent(index + 1)) >= 3D,
                    "Slice " + index + " label contrast is below 3:1.");
            }

            Assert.Empty(presentation.ValidateDocument());
        }
    
        [Fact]
        public void TemplateDeck_KeepsTheTemplateTitleAndStartsContentBelowIt() {
            byte[] template;
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                source.AddSlide();
                template = source.ToBytes();
            }

            using var input = new MemoryStream(template);
            using PowerPointPresentation presentation = PowerPointPresentation.Load(input);
            var slide = new PowerPointLayoutSlide { Title = "Quarter at a glance" };
            slide.Blocks.Add(new PowerPointLayoutKpi("$4.2M", "Revenue", "+12%"));
            slide.Blocks.Add(new PowerPointLayoutKpi("38%", "Margin", "+2 pt"));
            PowerPointSlide added = PowerPointLayoutEngine.AddSlides(presentation, slide).Single();

            P.Shape title = Shapes(added).Single(s => Placeholder(s)?.Type?.Value == P.PlaceholderValues.Title);
            Assert.Null(title.ShapeProperties!.Transform2D); // position, font and size come from the template layout
            Assert.Null(title.Descendants<A.RunProperties>().First().FontSize);
            double titleBottom = added.GetLayoutPlaceholder(PowerPointPlaceholderType.Title)!.Value.Bounds!.Value.BottomPoints;
            foreach (P.Shape shape in Shapes(added).Where(s => s.NonVisualShapeProperties!.NonVisualDrawingProperties!.Name!.Value!.StartsWith("Metric"))) {
                Assert.True(Points(shape.ShapeProperties!.Transform2D!.Offset!.Y!.Value) >= titleBottom, "Content overlaps the template title.");
            }
        }

        [Fact]
        public void Compose_ReportsLayoutDiagnosticsInThePreflight() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            var plan = new PowerPointDeckPlan().AddLogoWall("Partners", null, new[] { new PowerPointLogoItem("Contoso") },
                configure: options => options.FeaturedImagePath = "certificate.png");
            PowerPointCompositionResult result = presentation.Compose(plan,
                PowerPointCompositionOptions.FromDesign(PowerPointDeckDesign.FromBrand("#2463EB", "diagnostics")));
            Assert.Contains(result.Preflight.Findings, f => f.Code == "Layout.designer-option-ignored");
        }
    }
}
