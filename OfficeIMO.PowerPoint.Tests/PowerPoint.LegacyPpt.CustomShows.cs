using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Capabilities;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.Tests {
    public partial class PowerPointLegacyPptInteractionTests {
        [Fact]
        public void NativeWriter_AuthorsAndProjectsCustomShowAndAction() {
            byte[] bytes;
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                PowerPointSlide first = source.AddSlide(P.SlideLayoutValues.Blank);
                PowerPointSlide actionSlide = source.AddSlide(P.SlideLayoutValues.Blank);
                PowerPointSlide third = source.AddSlide(P.SlideLayoutValues.Blank);
                first.AddTextBox("First");
                third.AddTextBox("Third");
                AddCustomShow(source, 42, "Executive path", third, first);
                PowerPointAutoShape actionShape = actionSlide.AddRectangle(
                    100000, 100000, 1000000, 500000);
                GetDrawingProperties(actionShape).Append(new A.HyperlinkOnClick {
                    Id = string.Empty,
                    Action = "ppaction://customshow?id=42&return=true",
                    HighlightClick = true,
                    EndSound = true
                });

                LegacyPptWritePreflightReport report = source.AnalyzeLegacyPptWrite();
                Assert.True(report.CanWrite,
                    string.Join(Environment.NewLine, report.Findings));
                bytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation legacy = LegacyPptPresentation.Load(bytes);
            LegacyPptCustomShow show = Assert.Single(legacy.CustomShows);
            Assert.Equal("Executive path", show.Name);
            Assert.Equal(new[] { legacy.Slides[2].SlideId,
                legacy.Slides[0].SlideId }, show.SlideIds);
            Assert.False(show.HasUnresolvedSlides);
            LegacyPptInteraction action = Assert.Single(legacy.Slides[1].Shapes
                .SelectMany(shape => shape.Interactions));
            Assert.Equal(LegacyPptInteractionAction.CustomShow, action.Action);
            Assert.Equal("Executive path", action.Name);
            Assert.Same(show, action.CustomShow);
            Assert.True(action.IsAnimated);
            Assert.True(action.StopsSound);
            Assert.True(action.ReturnsFromCustomShow);
            LegacyPptImportReport inventory = legacy.CreateImportReport();
            Assert.Equal(1, inventory.CustomShowCount);
            Assert.Equal(2, inventory.CustomShowSlideEntryCount);

            using var input = new MemoryStream(bytes, writable: false);
            using PowerPointPresentation projected = PowerPointPresentation.Load(input);
            P.CustomShow projectedShow = Assert.Single(projected.OpenXmlDocument
                .PresentationPart!.Presentation!.CustomShowList!
                .Elements<P.CustomShow>());
            Assert.Equal("Executive path", projectedShow.Name?.Value);
            AssertCustomShowSlides(projected, projectedShow,
                projected.Slides[2], projected.Slides[0]);
            A.HyperlinkOnClick projectedAction = projected.Slides[1].SlidePart
                .Slide!.Descendants<A.HyperlinkOnClick>().Single();
            Assert.Equal("ppaction://customshow?id=" + projectedShow.Id!.Value
                + "&return=true", projectedAction.Action?.Value);
            Assert.True(projectedAction.HighlightClick?.Value);
            Assert.True(projectedAction.EndSound?.Value);
            Assert.Empty(projected.ValidateDocument());
            Assert.Equal(bytes, projected.ToBytes(PowerPointFileFormat.Ppt));
        }

        [Fact]
        public void ImportedCustomShow_AddEditAndRemoveAppendsPreservingRecords() {
            byte[] sourceBytes;
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                source.AddSlide(P.SlideLayoutValues.Blank).AddRectangle(
                    100000, 100000, 1000000, 500000);
                source.AddSlide(P.SlideLayoutValues.Blank).AddTextBox("Destination");
                sourceBytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation original = LegacyPptPresentation.Load(sourceBytes);

            byte[] addedBytes;
            using (var input = new MemoryStream(sourceBytes, writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation.Load(input)) {
                AddCustomShow(imported, 99, "Short path", imported.Slides[1]);
                GetDrawingProperties(Assert.Single(imported.Slides[0].Shapes))
                    .Append(new A.HyperlinkOnClick {
                        Id = string.Empty,
                        Action = "ppaction://customshow?id=99"
                    });
                LegacyPptWritePreflightReport report = imported.AnalyzeLegacyPptWrite();
                Assert.True(report.CanWrite,
                    string.Join(Environment.NewLine, report.Findings));
                addedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation added = LegacyPptPresentation.Load(addedBytes);
            LegacyPptCustomShow addedShow = Assert.Single(added.CustomShows);
            Assert.Equal("Short path", addedShow.Name);
            Assert.Equal(new[] { added.Slides[1].SlideId }, addedShow.SlideIds);
            Assert.Equal("Short path", Assert.Single(added.Slides[0].Shapes
                .SelectMany(shape => shape.Interactions)).Name);
            Assert.True(added.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));

            byte[] editedBytes;
            using (var input = new MemoryStream(addedBytes, writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation.Load(input)) {
                P.CustomShow show = imported.OpenXmlDocument.PresentationPart!
                    .Presentation!.CustomShowList!.Elements<P.CustomShow>().Single();
                show.Name = "Full path";
                show.SlideList!.RemoveAllChildren<P.SlideListEntry>();
                PresentationPart presentationPart = imported.OpenXmlDocument
                    .PresentationPart;
                show.SlideList.Append(
                    new P.SlideListEntry {
                        Id = presentationPart.GetIdOfPart(
                            imported.Slides[0].SlidePart)
                    },
                    new P.SlideListEntry {
                        Id = presentationPart.GetIdOfPart(
                            imported.Slides[1].SlidePart)
                    });
                A.HyperlinkOnClick action = imported.Slides[0].SlidePart.Slide!
                    .Descendants<A.HyperlinkOnClick>().Single();
                action.Action = "ppaction://customshow?id=" + show.Id!.Value
                    + "&return=true";
                Assert.Equal(string.Empty, action.Id?.Value);
                Assert.True(string.IsNullOrEmpty(action.Tooltip?.Value));
                Assert.Single(imported.OpenXmlDocument.PresentationPart!
                    .Presentation!.CustomShowList!.Elements<P.CustomShow>(),
                    candidate => candidate.Id?.Value == show.Id.Value);
                LegacyPptWritePreflightReport report = imported
                    .AnalyzeLegacyPptWrite();
                Assert.True(report.CanWrite,
                    string.Join(Environment.NewLine, report.Findings));
                editedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation edited = LegacyPptPresentation.Load(editedBytes);
            LegacyPptCustomShow editedShow = Assert.Single(edited.CustomShows);
            Assert.Equal("Full path", editedShow.Name);
            Assert.Equal(new[] { edited.Slides[0].SlideId,
                edited.Slides[1].SlideId }, editedShow.SlideIds);
            LegacyPptInteraction editedAction = Assert.Single(edited.Slides[0]
                .Shapes.SelectMany(shape => shape.Interactions));
            Assert.Equal("Full path", editedAction.Name);
            Assert.True(editedAction.ReturnsFromCustomShow);
            Assert.True(edited.Package.DocumentStream.AsSpan(0,
                    added.Package.DocumentStream.Length)
                .SequenceEqual(added.Package.DocumentStream));

            byte[] removedBytes;
            using (var input = new MemoryStream(editedBytes, writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation.Load(input)) {
                imported.OpenXmlDocument.PresentationPart!.Presentation!
                    .CustomShowList = null;
                GetDrawingProperties(Assert.Single(imported.Slides[0].Shapes))
                    .RemoveAllChildren<A.HyperlinkOnClick>();
                Assert.True(imported.AnalyzeLegacyPptWrite().CanWrite);
                removedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation removed = LegacyPptPresentation.Load(removedBytes);
            Assert.Empty(removed.CustomShows);
            Assert.Empty(removed.Slides[0].Shapes
                .SelectMany(shape => shape.Interactions));
            Assert.True(removed.Package.DocumentStream.AsSpan(0,
                    edited.Package.DocumentStream.Length)
                .SequenceEqual(edited.Package.DocumentStream));
        }

        [Fact]
        public void NativeWriter_BlocksAmbiguousCustomShowNames() {
            using PowerPointPresentation source = PowerPointPresentation.Create();
            PowerPointSlide slide = source.AddSlide(P.SlideLayoutValues.Blank);
            AddCustomShow(source, 1, "Duplicate", slide);
            P.CustomShow duplicate = new(new P.SlideList(
                new P.SlideListEntry {
                    Id = source.OpenXmlDocument.PresentationPart!
                        .GetIdOfPart(slide.SlidePart)
                })) {
                Id = 2,
                Name = "Duplicate"
            };
            source.OpenXmlDocument.PresentationPart!.Presentation!
                .CustomShowList!.Append(duplicate);

            LegacyPptWriteFinding finding = Assert.Single(
                source.AnalyzeLegacyPptWrite().Findings,
                item => item.Code == "PPT-WRITE-CUSTOM-SHOW");
            Assert.Contains("unique", finding.Description,
                StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void NativeWriter_BlocksMultipleSlideListsInCustomShow() {
            using PowerPointPresentation source =
                PowerPointPresentation.Create();
            PowerPointSlide slide = source.AddSlide(
                P.SlideLayoutValues.Blank);
            AddCustomShow(source, 1, "Duplicated list", slide);
            P.CustomShow show = source.OpenXmlDocument.PresentationPart!
                .Presentation!.CustomShowList!.Elements<P.CustomShow>()
                .Single();
            show.Append(new P.SlideList(new P.SlideListEntry {
                Id = source.OpenXmlDocument.PresentationPart
                    .GetIdOfPart(slide.SlidePart)
            }));

            LegacyPptWriteFinding finding = Assert.Single(
                source.AnalyzeLegacyPptWrite().Findings,
                item => item.Code == "PPT-WRITE-CUSTOM-SHOW");

            Assert.Contains("extension data", finding.Description,
                StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void ImportedCustomShow_TracksAppendedAndRemovedSlides() {
            byte[] sourceBytes;
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                PowerPointSlide first = source.AddSlide(P.SlideLayoutValues.Blank);
                first.AddTextBox("First");
                AddCustomShow(source, 7, "Mutable path", first);
                sourceBytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation original = LegacyPptPresentation.Load(sourceBytes);

            byte[] appendedBytes;
            using (var input = new MemoryStream(sourceBytes, writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation.Load(input)) {
                PowerPointSlide appended = imported.AddSlide(P.SlideLayoutValues.Blank);
                appended.AddTextBox("Appended");
                P.CustomShow show = imported.OpenXmlDocument.PresentationPart!
                    .Presentation!.CustomShowList!.Elements<P.CustomShow>().Single();
                show.SlideList!.Append(new P.SlideListEntry {
                    Id = imported.OpenXmlDocument.PresentationPart
                        .GetIdOfPart(appended.SlidePart)
                });
                LegacyPptWritePreflightReport report = imported
                    .AnalyzeLegacyPptWrite();
                Assert.True(report.CanWrite,
                    string.Join(Environment.NewLine, report.Findings));
                appendedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation withAppended = LegacyPptPresentation.Load(
                appendedBytes);
            Assert.Equal(new[] { withAppended.Slides[0].SlideId,
                withAppended.Slides[1].SlideId },
                Assert.Single(withAppended.CustomShows).SlideIds);
            Assert.True(withAppended.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));

            byte[] removedBytes;
            using (var input = new MemoryStream(appendedBytes, writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation.Load(input)) {
                imported.RemoveSlide(1);
                LegacyPptWritePreflightReport report = imported
                    .AnalyzeLegacyPptWrite();
                Assert.True(report.CanWrite,
                    string.Join(Environment.NewLine, report.Findings));
                removedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation removed = LegacyPptPresentation.Load(removedBytes);
            Assert.Single(removed.Slides);
            Assert.Equal(new[] { removed.Slides[0].SlideId },
                Assert.Single(removed.CustomShows).SlideIds);
            Assert.True(removed.Package.DocumentStream.AsSpan(0,
                    withAppended.Package.DocumentStream.Length)
                .SequenceEqual(withAppended.Package.DocumentStream));
        }

        private static void AddCustomShow(PowerPointPresentation presentation,
            uint id, string name, params PowerPointSlide[] slides) {
            PresentationPart presentationPart = presentation.OpenXmlDocument
                .PresentationPart!;
            var slideList = new P.SlideList(slides.Select(slide =>
                new P.SlideListEntry {
                    Id = presentationPart.GetIdOfPart(slide.SlidePart)
                }));
            var show = new P.CustomShow(slideList) { Id = id, Name = name };
            presentationPart.Presentation!.CustomShowList ??= new P.CustomShowList();
            presentationPart.Presentation.CustomShowList.Append(show);
        }

        private static void AssertCustomShowSlides(
            PowerPointPresentation presentation, P.CustomShow show,
            params PowerPointSlide[] expected) {
            PresentationPart presentationPart = presentation.OpenXmlDocument
                .PresentationPart!;
            SlidePart[] actual = show.SlideList!.Elements<P.SlideListEntry>()
                .Select(entry => {
                    Assert.True(presentationPart.TryGetPartById(entry.Id!.Value!,
                        out OpenXmlPart? part));
                    return Assert.IsType<SlidePart>(part);
                }).ToArray();
            Assert.Equal(expected.Select(slide => slide.SlidePart), actual);
        }
    }
}