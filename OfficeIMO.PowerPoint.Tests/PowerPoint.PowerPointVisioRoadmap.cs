using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Capabilities;
using OfficeIMO.PowerPoint.LegacyPpt.Write;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Dgm = DocumentFormat.OpenXml.Drawing.Diagrams;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Tests {
    public class PowerPointPowerPointVisioRoadmapTests {
        [Theory]
        [InlineData("bar3DChart", OfficeChartKind.BarStacked100)]
        [InlineData("line3DChart", OfficeChartKind.LineStacked)]
        [InlineData("area3DChart", OfficeChartKind.AreaStacked100)]
        [InlineData("pie3DChart", OfficeChartKind.Pie)]
        [InlineData("ofPieChart", OfficeChartKind.Pie)]
        [InlineData("stockChart", OfficeChartKind.Line)]
        [InlineData("surfaceChart", OfficeChartKind.Line)]
        [InlineData("surface3DChart", OfficeChartKind.Line)]
        public void AdvancedChartProjectionRetainsRepresentableFamilySemantics(
            string family, OfficeChartKind expected) {
            OpenXmlElement group = family switch {
                "bar3DChart" => new C.Bar3DChart(
                    new C.BarDirection { Val = C.BarDirectionValues.Bar },
                    new C.BarGrouping { Val = C.BarGroupingValues.PercentStacked }),
                "line3DChart" => new C.Line3DChart(
                    new C.Grouping { Val = C.GroupingValues.Stacked }),
                "area3DChart" => new C.Area3DChart(
                    new C.Grouping { Val = C.GroupingValues.PercentStacked }),
                "pie3DChart" => new C.Pie3DChart(),
                "ofPieChart" => new C.OfPieChart(),
                "stockChart" => new C.StockChart(),
                "surfaceChart" => new C.SurfaceChart(),
                "surface3DChart" => new C.Surface3DChart(),
                _ => throw new InvalidOperationException()
            };
            Assert.Equal(expected, PowerPointChart.GetAdvancedProjection(group));
        }

        [Fact]
        public void AdvancedImportedChartEditsCachesWithoutReplacingNativeFamily() {
            using var stream = new MemoryStream();
            using (PowerPointPresentation presentation = PowerPointPresentation.Create(stream)) {
                PowerPointChart chart = presentation.AddSlide().AddChart();
                ChartPart part = presentation.Slides[0].SlidePart.ChartParts.Single();
                C.PlotArea plot = part.ChartSpace!.GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!;
                C.BarChart bar = plot.GetFirstChild<C.BarChart>()!;
                var advanced = new C.Bar3DChart();
                foreach (var child in bar.ChildElements.ToArray()) advanced.Append(child.CloneNode(true));
                advanced.GetFirstChild<C.BarDirection>()!.Val = C.BarDirectionValues.Bar;
                advanced.GetFirstChild<C.BarGrouping>()!.Val = C.BarGroupingValues.PercentStacked;
                plot.ReplaceChild(advanced, bar);

                PowerPointImportedChartReport report = chart.InspectImportedContent();
                Assert.Equal("bar3DChart", report.Family);
                Assert.Equal(PowerPointImportedChartSupport.EditableWithProjectedRendering, report.Support);
                Assert.Equal(OfficeChartKind.BarStacked100, report.ExportProjection);

                Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot before));
                OfficeChartSeries[] replacementSeries = before.Data.Series
                    .Select((series, index) => new OfficeChartSeries(
                        index == 0 ? "Revenue" : series.Name,
                        new[] { 10D + index, 20D + index })).ToArray();
                chart.UpdateData(new OfficeChartData(new[] { "North", "South" },
                    replacementSeries));
                Assert.NotNull(plot.GetFirstChild<C.Bar3DChart>());
                Assert.Null(plot.GetFirstChild<C.BarChart>());
                Assert.Equal("Revenue", advanced.Descendants<C.StringPoint>().First().NumericValue!.Text);
                Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
                Assert.Equal(OfficeChartKind.BarStacked100, snapshot.ChartKind);
                Assert.Contains("semantic projection", report.Diagnostics.Single(), StringComparison.OrdinalIgnoreCase);
                OfficeImageExportResult svg = presentation.Slides[0]
                    .ExportImage(OfficeImageExportFormat.Svg);
                Assert.NotEmpty(svg.Bytes);
                Assert.Contains("Revenue", System.Text.Encoding.UTF8.GetString(svg.Bytes),
                    StringComparison.Ordinal);
                byte[] pdf = presentation.ToPdf();
                Assert.True(pdf.Length > 100);
                Assert.Equal("%PDF-", System.Text.Encoding.ASCII.GetString(pdf, 0, 5));
            }
        }

        [Fact]
        public void SmartArtTopologyCanReorderProcessButRejectsMeaningChangingParent() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            PowerPointSmartArt smartArt = presentation.AddSlide().AddSmartArt(
                PowerPointSmartArtType.BasicProcess, new[] { "Plan", "Build", "Ship" });
            Assert.True(smartArt.TryGetTopology(out PowerPointSmartArtTopology topology, out string diagnostic), diagnostic);
            PowerPointSmartArtNode[] reordered = topology.Nodes.Reverse().ToArray();
            for (uint index = 0; index < reordered.Length; index++) reordered[index].Order = index;
            smartArt.UpdateTopology(reordered);
            Assert.Equal(new[] { "Ship", "Build", "Plan" }, smartArt.GetNodeTexts());
            reordered[0].ParentId = reordered[1].Id;
            Assert.Throws<InvalidOperationException>(() => smartArt.UpdateTopology(reordered));
        }

        [Fact]
        public void AdvancedChartWithoutEditableWorkbookIsDiagnosedAndLeftUnchanged() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            PowerPointChart chart = presentation.AddSlide().AddChart();
            ChartPart part = presentation.Slides[0].SlidePart.ChartParts.Single();
            C.PlotArea plot = part.ChartSpace!.GetFirstChild<C.Chart>()!
                .GetFirstChild<C.PlotArea>()!;
            C.BarChart bar = plot.GetFirstChild<C.BarChart>()!;
            var advanced = new C.Bar3DChart();
            foreach (OpenXmlElement child in bar.ChildElements.ToArray())
                advanced.Append(child.CloneNode(true));
            plot.ReplaceChild(advanced, bar);
            EmbeddedPackagePart embedded = part.GetPartsOfType<EmbeddedPackagePart>().Single();
            part.DeletePart(embedded);

            PowerPointImportedChartReport report = chart.InspectImportedContent();
            Assert.Equal(PowerPointImportedChartSupport.PreservationOnly, report.Support);
            Assert.Contains(report.Diagnostics, detail => detail.Contains(
                "no single referenced embedded workbook",
                StringComparison.OrdinalIgnoreCase));
            Assert.Throws<NotSupportedException>(() => chart.UpdateData(
                new OfficeChartData(new[] { "A" },
                    new[] { new OfficeChartSeries("S", new[] { 1D }) })));
            Assert.Same(advanced, plot.GetFirstChild<C.Bar3DChart>());
        }

        [Fact]
        public void AdvancedChartRejectsRicherReferencedWorkbookWithoutMutation() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            PowerPointChart chart = presentation.AddSlide().AddChart();
            ChartPart part = presentation.Slides[0].SlidePart.ChartParts.Single();
            C.PlotArea plot = part.ChartSpace!.GetFirstChild<C.Chart>()!
                .GetFirstChild<C.PlotArea>()!;
            C.BarChart bar = plot.GetFirstChild<C.BarChart>()!;
            var advanced = new C.Bar3DChart();
            foreach (OpenXmlElement child in bar.ChildElements.ToArray())
                advanced.Append(child.CloneNode(true));
            plot.ReplaceChild(advanced, bar);
            EmbeddedPackagePart embedded = part.GetPartsOfType<EmbeddedPackagePart>().Single();
            var workbookBytes = new MemoryStream();
            using (Stream input = embedded.GetStream(FileMode.Open, FileAccess.Read))
                input.CopyTo(workbookBytes);
            using (SpreadsheetDocument workbook = SpreadsheetDocument.Open(
                       workbookBytes, true)) {
                WorksheetPart extra = workbook.WorkbookPart!.AddNewPart<WorksheetPart>();
                extra.Worksheet = new S.Worksheet(new S.SheetData());
                workbook.WorkbookPart.Workbook.Sheets!.Append(new S.Sheet {
                    Id = workbook.WorkbookPart.GetIdOfPart(extra),
                    Name = "Producer Details", SheetId = 2U
                });
                workbook.WorkbookPart.Workbook.Save();
            }
            byte[] richer = workbookBytes.ToArray();
            using (var input = new MemoryStream(richer)) embedded.FeedData(input);

            PowerPointImportedChartReport report = chart.InspectImportedContent();
            Assert.Equal(PowerPointImportedChartSupport.PreservationOnly,
                report.Support);
            Assert.Contains("richer", report.Diagnostics.Single(),
                StringComparison.OrdinalIgnoreCase);
            Assert.Throws<NotSupportedException>(() => chart.UpdateData(
                new OfficeChartData(new[] { "A" }, new[] {
                    new OfficeChartSeries("S", new[] { 1D })
                })));
            using Stream after = embedded.GetStream(FileMode.Open, FileAccess.Read);
            using var saved = new MemoryStream();
            after.CopyTo(saved);
            Assert.Equal(richer, saved.ToArray());
        }

        [Fact]
        public void MixedChartWithAdvancedGroupIsPreservationOnlyInsteadOfPartialProjection() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            PowerPointChart chart = presentation.AddSlide().AddChart();
            ChartPart part = presentation.Slides[0].SlidePart.ChartParts.Single();
            C.PlotArea plot = part.ChartSpace!.GetFirstChild<C.Chart>()!
                .GetFirstChild<C.PlotArea>()!;
            C.BarChart bar = plot.GetFirstChild<C.BarChart>()!;
            var line = new C.LineChart(new C.Grouping {
                Val = C.GroupingValues.Standard
            });
            foreach (C.BarChartSeries source in bar.Elements<C.BarChartSeries>())
                line.Append(new C.LineChartSeries(source.ChildElements
                    .Select(child => child.CloneNode(true))));
            var advanced = new C.StockChart();
            foreach (C.BarChartSeries source in bar.Elements<C.BarChartSeries>())
                advanced.Append(new C.LineChartSeries(source.ChildElements
                    .Select(child => child.CloneNode(true))));
            plot.InsertAfter(line, bar);
            plot.InsertAfter(advanced, line);

            PowerPointImportedChartReport report = chart.InspectImportedContent();
            Assert.Equal("mixed", report.Family);
            Assert.Equal(PowerPointImportedChartSupport.PreservationOnly,
                report.Support);
            Assert.False(chart.TryGetOfficeSnapshot(out _));
        }

        [Fact]
        public void TimelineEditsMotionRotationAndCommandWithoutRemovingUnknownSibling() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointTextBox shape = slide.AddTextBox("Animate me");
            var unknown = new AnimateColor();
            slide.SlidePart.Slide!.Timing = new Timing(new TimeNodeList(
                new ParallelTimeNode(new CommonTimeNode(new ChildTimeNodeList(unknown)) {
                    Id = 1U, Duration = "indefinite", NodeType = TimeNodeValues.TmingRoot
                })));
            PowerPointTimelineAction motion = slide.AddMotionAnimation(shape, "M 0 0 L 1 0 E");
            PowerPointTimelineAction rotation = slide.AddRotationAnimation(shape, 90D);
            PowerPointTimelineAction command = slide.AddCommandAnimation(shape, "play");
            Assert.True(slide.SetAnimationDuration(rotation.TimingId, 750U));
            Assert.True(slide.RemoveAnimation(motion.TimingId));
            Assert.Single(slide.SlidePart.Slide.Timing.Descendants<AnimateColor>());
            Assert.Single(slide.SlidePart.Slide.Timing.Descendants<AnimateRotation>());
            Assert.Single(slide.SlidePart.Slide.Timing.Descendants<Command>());
            Assert.Equal(750U.ToString(), slide.SlidePart.Slide.Timing.Descendants<AnimateRotation>()
                .Single().CommonBehavior!.CommonTimeNode!.Duration!.Value);
            Assert.Equal(PowerPointAnimationKind.Command, command.Kind);

            PowerPointSlide emptyTimelineSlide = presentation.AddSlide();
            PowerPointTextBox emptyTimelineShape = emptyTimelineSlide.AddTextBox("Fresh");
            emptyTimelineSlide.AddMotionAnimation(emptyTimelineShape,
                "M 0 0 L 0 1 E");
            uint[] timingIds = emptyTimelineSlide.SlidePart.Slide!.Timing!
                .Descendants<CommonTimeNode>().Select(node => node.Id!.Value).ToArray();
            Assert.Equal(timingIds.Length, timingIds.Distinct().Count());
        }

        [Fact]
        public void RemovingTypedAnimationPreservesUnknownSiblingInsideItsOwner() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointTimelineAction motion = slide.AddMotionAnimation(
                slide.AddTextBox("Target"), "M 0 0 L 1 0 E");
            AnimateMotion action = slide.SlidePart.Slide!.Timing!
                .Descendants<AnimateMotion>().Single();
            ChildTimeNodeList ownerList = (ChildTimeNodeList)action.Parent!;
            var producer = new OpenXmlUnknownElement("p99", "producer",
                "urn:producer:timing");
            producer.SetAttribute(new OpenXmlAttribute("", "value", "", "keep"));
            ownerList.Append(producer);

            Assert.True(slide.RemoveAnimation(motion.TimingId));
            Assert.Single(slide.SlidePart.Slide.Timing!
                .Descendants<OpenXmlUnknownElement>());
            Assert.Contains("urn:producer:timing",
                slide.SlidePart.Slide.Timing!.OuterXml, StringComparison.Ordinal);
        }

        [Fact]
        public void SmartArtRejectsCustomLayoutAndPreservesRichTextOnTopologyOnlyEdit() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointSmartArt smartArt = slide.AddSmartArt(
                PowerPointSmartArtType.BasicProcess, new[] { "Plan", "Build" });
            DiagramDataPart dataPart = slide.SlidePart.DiagramDataParts.Single();
            XDocument data;
            using (Stream input = dataPart.GetStream(FileMode.Open, FileAccess.Read))
                data = XDocument.Load(input);
            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
            XElement planText = data.Descendants(a + "t")
                .First(element => element.Value == "Plan");
            planText.Parent!.AddBeforeSelf(new XElement(a + "r",
                new XElement(a + "rPr", new XAttribute("b", "1")),
                new XElement(a + "t", "Rich ")));
            using (Stream output = dataPart.GetStream(FileMode.Create, FileAccess.Write))
                data.Save(output);

            Assert.True(smartArt.TryGetTopology(out PowerPointSmartArtTopology topology,
                out string diagnostic), diagnostic);
            PowerPointSmartArtNode[] reordered = topology.Nodes.Reverse().ToArray();
            for (uint index = 0; index < reordered.Length; index++)
                reordered[index].Order = index;
            smartArt.UpdateTopology(reordered);
            Assert.Contains("b=\"1\"", dataPart.DataModelRoot!.OuterXml,
                StringComparison.Ordinal);
            PowerPointSmartArtNode richNode = reordered.Single(node =>
                node.Text == "Rich Plan");
            richNode.Text = "Changed";
            Assert.Throws<NotSupportedException>(() =>
                smartArt.UpdateTopology(reordered));
            Assert.Contains("Rich ", dataPart.DataModelRoot!.OuterXml,
                StringComparison.Ordinal);

            DiagramLayoutDefinitionPart layout = slide.SlidePart
                .DiagramLayoutDefinitionParts.Single();
            layout.LayoutDefinition!.SetAttribute(new OpenXmlAttribute(
                "producer", "flag", "urn:producer:smartart", "1"));
            Assert.False(smartArt.TryGetTopology(out _, out diagnostic));
            Assert.Contains("canonical", diagnostic,
                StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void PlaybackValidationAndLinkedOlePreflightAreNonDestructive() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointMedia video = slide.AddLinkedVideo(
                new Uri("https://example.test/video.mp4"));
            ((DocumentFormat.OpenXml.Presentation.Picture)video.Element)
                .Descendants<P14.Media>().Single().Remove();
            string timingBefore = slide.SlidePart.Slide!.Timing!.OuterXml;
            Assert.Throws<InvalidOperationException>(() => video.SetPlaybackOptions(
                new PowerPointMediaPlaybackOptions { VolumePercent = 12,
                    Mute = true, Loop = true }));
            Assert.Equal(timingBefore, slide.SlidePart.Slide.Timing!.OuterXml);

            slide.AddLinkedOleObject(new Uri("https://example.test/data.xlsx"),
                "Excel.Sheet.12");
            LegacyPptWritePreflightReport preflight = presentation
                .AnalyzeLegacyPptWrite();
            Assert.Contains(preflight.Findings, finding =>
                finding.Feature == LegacyPptFeature.LinkedOle &&
                finding.Code == "PPT-WRITE-LINKED-OLE");
        }

        [Fact]
        public void LinkedMediaAndOleRoundTripWithTypedPlaybackAndTargets() {
            string path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(path)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointMedia video = slide.AddLinkedVideo(new Uri("https://example.test/video.mp4"));
                    video.SetPlaybackOptions(new PowerPointMediaPlaybackOptions {
                        VolumePercent = 55, Mute = true, Loop = true,
                        FullScreen = true, TrimStartMilliseconds = 1000,
                        TrimEndMilliseconds = 9000, FadeInMilliseconds = 250
                    });
                    PowerPointOleObject ole = slide.AddLinkedOleObject(
                        new Uri("https://example.test/data.xlsx"), "Excel.Sheet.12", true);
                    presentation.Save();
                    Assert.Equal(PowerPointMediaSourceKind.Linked, video.SourceKind);
                    Assert.True(ole.IsLinked);
                }
                using PowerPointPresentation loaded = PowerPointPresentation.Load(path);
                PowerPointMedia loadedVideo = Assert.Single(loaded.Slides[0].Media);
                Assert.Equal(new Uri("https://example.test/video.mp4"), loadedVideo.LinkUri);
                PowerPointMediaPlaybackOptions playback = loadedVideo.GetPlaybackOptions();
                Assert.Equal(55, playback.VolumePercent);
                Assert.True(playback.Mute);
                Assert.True(playback.Loop);
                Assert.True(playback.FullScreen);
                Assert.Equal(1000U, playback.TrimStartMilliseconds);
                PowerPointOleObject loadedOle = Assert.Single(loaded.Slides[0].OleObjects);
                Assert.True(loadedOle.AutoUpdate);
                Assert.Equal(new Uri("https://example.test/data.xlsx"), loadedOle.LinkUri);
                loadedOle.UpdateLink(new Uri("https://example.test/new.xlsx"));
                Assert.Equal(new Uri("https://example.test/new.xlsx"), loadedOle.LinkUri);
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }
    }
}
