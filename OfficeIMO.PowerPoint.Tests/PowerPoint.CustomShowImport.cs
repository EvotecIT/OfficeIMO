using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests {
    public sealed class PowerPointCustomShowImportTests {
        [Fact]
        public void ImportRemapsReferencedCustomShowAndExportRemovesIt() {
            using PowerPointPresentation source =
                PowerPointPresentation.Create();
            PowerPointSlide requested = source.AddSlide();
            PowerPointAutoShape actionShape = requested.AddRectangle(
                100000, 100000, 1000000, 500000);
            PowerPointSlide firstShowSlide = source.AddSlide();
            firstShowSlide.AddTitle("First show slide");
            PowerPointSlide secondShowSlide = source.AddSlide();
            secondShowSlide.AddTitle("Second show slide");
            PresentationPart sourcePart = source.OpenXmlDocument
                .PresentationPart!;
            var sourceShow = new CustomShow(new SlideList(
                    new SlideListEntry {
                        Id = sourcePart.GetIdOfPart(firstShowSlide.SlidePart)
                    },
                    new SlideListEntry {
                        Id = sourcePart.GetIdOfPart(secondShowSlide.SlidePart)
                    })) {
                    Id = 17U,
                    Name = "Tour"
                };
            var extension = new Extension {
                Uri = "{40A09A7A-19E1-4D9D-A417-7F2234A3D10B}"
            };
            extension.Append(new OpenXmlUnknownElement("p15",
                "customShowMarker",
                "http://schemas.microsoft.com/office/powerpoint/2012/main"));
            sourceShow.Append(new ExtensionList(extension));
            sourcePart.Presentation!.CustomShowList =
                new CustomShowList(sourceShow);
            NonVisualDrawingProperties actionProperties =
                ((Shape)actionShape.Element).NonVisualShapeProperties!
                .NonVisualDrawingProperties!;
            actionProperties.Append(new A.HyperlinkOnClick {
                Id = string.Empty,
                Action = "ppaction://customshow?id=17&return=true"
            });

            using PowerPointPresentation target =
                PowerPointPresentation.Create();
            PowerPointSlide existingTarget = target.AddSlide();
            PresentationPart targetPart = target.OpenXmlDocument
                .PresentationPart!;
            targetPart.Presentation!.CustomShowList = new CustomShowList(
                new CustomShow(new SlideList(new SlideListEntry {
                    Id = targetPart.GetIdOfPart(existingTarget.SlidePart)
                })) {
                    Id = 17U,
                    Name = "Tour"
                });

            PowerPointSlide imported = target.ImportSlide(source, 0);

            Assert.Equal(4, target.Slides.Count);
            CustomShow[] shows = targetPart.Presentation.CustomShowList!
                .Elements<CustomShow>().ToArray();
            Assert.Equal(2, shows.Length);
            CustomShow importedShow = Assert.Single(shows,
                show => show.Id?.Value != 17U);
            Assert.Equal("Tour (2)", importedShow.Name?.Value);
            Assert.Equal(extension.OuterXml,
                importedShow.ExtensionList!.Elements<Extension>()
                    .Single().OuterXml);
            SlidePart[] importedShowSlides = importedShow.SlideList!
                .Elements<SlideListEntry>()
                .Select(entry => (SlidePart)targetPart.GetPartById(
                    entry.Id!.Value!))
                .ToArray();
            Assert.Equal(new[] {
                target.Slides[2].SlidePart,
                target.Slides[3].SlidePart
            }, importedShowSlides);
            string action = imported.SlidePart.Slide!
                .Descendants<A.HyperlinkOnClick>().Single()
                .Action!.Value!;
            Assert.Equal("ppaction://customshow?id=" +
                importedShow.Id!.Value + "&return=true", action);
            Assert.Empty(target.ValidateDocument());

            using var exportedBytes = new MemoryStream();
            source.ExportSlide(0, exportedBytes);
            exportedBytes.Position = 0;
            using PowerPointPresentation exported =
                PowerPointPresentation.Load(exportedBytes);
            Assert.Single(exported.Slides);
            Assert.Null(exported.OpenXmlDocument.PresentationPart!
                .Presentation!.CustomShowList);
            Assert.Empty(exported.Slides[0].SlidePart.Slide!
                .Descendants<A.HyperlinkOnClick>());
            Assert.Empty(exported.ValidateDocument());
        }

        [Fact]
        public void ImportRemapsCustomShowActionsInClonedLayoutsAndExportCleansThem() {
            using PowerPointPresentation source =
                PowerPointPresentation.Create();
            PowerPointSlide requested = source.AddSlide();
            SlideLayoutPart sourceLayout = requested.SlidePart
                .SlideLayoutPart!;
            sourceLayout.SlideLayout!.CommonSlideData!.Name =
                "Custom show import layout";
            PowerPointSlide showSlide = source.AddSlide();
            showSlide.AddTitle("Layout show target");
            PresentationPart sourcePart = source.OpenXmlDocument
                .PresentationPart!;
            sourcePart.Presentation!.CustomShowList = new CustomShowList(
                new CustomShow(new SlideList(new SlideListEntry {
                    Id = sourcePart.GetIdOfPart(showSlide.SlidePart)
                })) {
                    Id = 23U,
                    Name = "Layout tour"
                });
            NonVisualDrawingProperties layoutProperties = sourceLayout
                .SlideLayout.Descendants<NonVisualDrawingProperties>()
                .First();
            layoutProperties.Append(new A.HyperlinkOnClick {
                Id = string.Empty,
                Action = "ppaction://customshow?id=23&return=true"
            });
            SlideMasterPart sourceMaster = sourceLayout.SlideMasterPart!;
            NonVisualDrawingProperties masterProperties = sourceMaster
                .SlideMaster!.Descendants<NonVisualDrawingProperties>()
                .First();
            masterProperties.Append(new A.HyperlinkOnClick {
                Id = string.Empty,
                Action = "ppaction://customshow?id=23&return=true"
            });
            sourceLayout.SlideLayout.Save();
            sourceMaster.SlideMaster.Save();

            using PowerPointPresentation target =
                PowerPointPresentation.Create();
            PowerPointSlide existing = target.AddSlide();
            PresentationPart targetPart = target.OpenXmlDocument
                .PresentationPart!;
            targetPart.Presentation!.CustomShowList = new CustomShowList(
                new CustomShow(new SlideList(new SlideListEntry {
                    Id = targetPart.GetIdOfPart(existing.SlidePart)
                })) {
                    Id = 23U,
                    Name = "Layout tour"
                });

            PowerPointSlide imported = target.ImportSlide(source, 0);

            Assert.Equal(3, target.Slides.Count);
            CustomShow importedShow = Assert.Single(targetPart.Presentation
                .CustomShowList!.Elements<CustomShow>(),
                show => show.Id?.Value != 23U);
            A.HyperlinkOnClick importedLayoutAction = Assert.Single(imported
                .SlidePart.SlideLayoutPart!.SlideLayout!
                .Descendants<A.HyperlinkOnClick>());
            Assert.Equal("ppaction://customshow?id="
                + importedShow.Id!.Value + "&return=true",
                importedLayoutAction.Action?.Value);
            A.HyperlinkOnClick importedMasterAction = Assert.Single(imported
                .SlidePart.SlideLayoutPart.SlideMasterPart!.SlideMaster!
                .Descendants<A.HyperlinkOnClick>());
            Assert.Equal(importedLayoutAction.Action?.Value,
                importedMasterAction.Action?.Value);
            Assert.Empty(target.ValidateDocument());

            using var exportedBytes = new MemoryStream();
            source.ExportSlide(0, exportedBytes);
            exportedBytes.Position = 0;
            using PowerPointPresentation exported =
                PowerPointPresentation.Load(exportedBytes);
            PowerPointSlide exportedSlide = Assert.Single(exported.Slides);
            Assert.Null(exported.OpenXmlDocument.PresentationPart!
                .Presentation!.CustomShowList);
            Assert.Empty(exportedSlide.SlidePart.SlideLayoutPart!
                .SlideLayout!.Descendants<A.HyperlinkOnClick>());
            Assert.Empty(exportedSlide.SlidePart.SlideLayoutPart
                .SlideMasterPart!.SlideMaster!
                .Descendants<A.HyperlinkOnClick>());
            Assert.Empty(exported.ValidateDocument());
        }

        [Fact]
        public void ImportRemovesMalformedCustomShowRelationshipsWithoutExtraSlides() {
            using PowerPointPresentation source =
                PowerPointPresentation.Create();
            PowerPointSlide requested = source.AddSlide();
            PowerPointSlide unwantedTarget = source.AddSlide();
            PowerPointAutoShape externalShape = requested.AddRectangle(
                100000, 100000, 1000000, 500000);
            HyperlinkRelationship externalRelationship = requested
                .SlidePart.AddHyperlinkRelationship(
                    new Uri("https://example.test/discarded"),
                    isExternal: true);
            GetNonVisualProperties(externalShape).Append(
                new A.HyperlinkOnClick {
                    Id = externalRelationship.Id,
                    Action = "ppaction://customshow?id=not-a-number"
                });
            const string InternalRelationshipId = "rIdMalformedTarget";
            requested.SlidePart.AddPart(unwantedTarget.SlidePart,
                InternalRelationshipId);
            PowerPointAutoShape internalShape = requested.AddRectangle(
                100000, 700000, 1000000, 500000);
            GetNonVisualProperties(internalShape).Append(
                new A.HyperlinkOnClick {
                    Id = InternalRelationshipId,
                    Action = "ppaction://customshow?id=999"
                });

            using PowerPointPresentation target =
                PowerPointPresentation.Create();

            PowerPointSlide imported = target.ImportSlide(source, 0);

            Assert.Single(target.Slides);
            Assert.Empty(imported.SlidePart.Slide!
                .Descendants<A.HyperlinkOnClick>());
            Assert.Empty(imported.SlidePart.HyperlinkRelationships);
            Assert.DoesNotContain(imported.SlidePart.Parts,
                pair => pair.RelationshipId == InternalRelationshipId);
            Assert.Empty(target.ValidateDocument());
        }

        [Fact]
        public void ImportPreservesStructuralRelationshipsThatShareMalformedActionIds() {
            using PowerPointPresentation source =
                PowerPointPresentation.Create();
            PowerPointSlide requested = source.AddSlide();
            SlideLayoutPart sourceLayout = requested.SlidePart
                .SlideLayoutPart!;
            sourceLayout.SlideLayout!.CommonSlideData!.Name =
                "Structural relationship import layout";
            SlideMasterPart sourceMaster = sourceLayout.SlideMasterPart!;
            ThemePart sourceTheme = sourceMaster.ThemePart!;
            string themeRelationshipId = sourceMaster.GetIdOfPart(
                sourceTheme);
            sourceMaster.SlideMaster!.Descendants<
                    NonVisualDrawingProperties>().First()
                .Append(new A.HyperlinkOnClick {
                    Id = themeRelationshipId,
                    Action = "ppaction://customshow?id=invalid"
                });
            sourceMaster.SlideMaster.Save();

            requested.Notes.Text = "Preserved notes";
            NotesSlidePart sourceNotes = requested.SlidePart
                .NotesSlidePart!;
            const string BacklinkRelationshipId = "rIdNotesBacklink";
            sourceNotes.AddPart(requested.SlidePart,
                BacklinkRelationshipId);
            sourceNotes.NotesSlide!.Descendants<
                    NonVisualDrawingProperties>().First()
                .Append(new A.HyperlinkOnClick {
                    Id = BacklinkRelationshipId,
                    Action = "ppaction://customshow?id=invalid"
                });
            sourceNotes.NotesSlide.Save();

            using PowerPointPresentation target =
                PowerPointPresentation.Create();

            PowerPointSlide imported = target.ImportSlide(source, 0);

            SlideMasterPart importedMaster = imported.SlidePart
                .SlideLayoutPart!.SlideMasterPart!;
            Assert.NotNull(importedMaster.ThemePart);
            Assert.Empty(importedMaster.SlideMaster!
                .Descendants<A.HyperlinkOnClick>());
            NotesSlidePart importedNotes = imported.SlidePart
                .NotesSlidePart!;
            Assert.Contains(importedNotes.Parts, pair =>
                pair.RelationshipId == BacklinkRelationshipId
                && ReferenceEquals(pair.OpenXmlPart,
                    imported.SlidePart));
            Assert.Empty(importedNotes.NotesSlide!
                .Descendants<A.HyperlinkOnClick>());
            Assert.Equal("Preserved notes", imported.Notes.Text);
            Assert.Empty(target.ValidateDocument());
        }

        private static NonVisualDrawingProperties GetNonVisualProperties(
            PowerPointAutoShape shape) =>
            ((Shape)shape.Element).NonVisualShapeProperties!
                .NonVisualDrawingProperties!;
    }
}
