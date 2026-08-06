using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests {
    public sealed class PowerPointCustomShowTests {
        [Fact]
        public void CustomShows_CreateEditAndRoundTripThroughPublicApi() {
            string path = Path.Combine(Path.GetTempPath(),
                "OfficeIMO-CustomShows-" + Guid.NewGuid().ToString("N") + ".pptx");
            try {
                using (PowerPointPresentation presentation =
                       PowerPointPresentation.Create(path)) {
                    PowerPointSlide first = presentation.AddSlide();
                    first.AddTitle("First");
                    PowerPointSlide second = presentation.AddSlide();
                    second.AddTitle("Second");
                    PowerPointSlide third = presentation.AddSlide();
                    third.AddTitle("Third");

                    PowerPointCustomShow show = presentation.AddCustomShow(
                        "Executive path", new[] { third, first });

                    Assert.Equal(1U, show.Id);
                    Assert.Equal(new[] { third, first }, show.Slides);
                    Assert.Contains(presentation.InspectFeatures().EditableFeatures,
                        feature => feature.Name == "Custom shows" && feature.Count == 1);

                    show.InsertSlide(1, second).MoveSlide(2, 0);
                    Assert.True(show.RemoveSlide(second));
                    presentation.RenameCustomShow(show, "Decision path");
                    Assert.Equal(new[] { first, third }, show.Slides);
                    presentation.Save();
                }

                using PowerPointPresentation reopened =
                    PowerPointPresentation.Load(path);
                PowerPointCustomShow saved = Assert.Single(reopened.CustomShows);
                Assert.Equal("Decision path", saved.Name);
                Assert.Equal(new[] { reopened.Slides[0], reopened.Slides[2] },
                    saved.Slides);
                Assert.Equal(saved.Id,
                    reopened.GetCustomShow("decision PATH")!.Id);

                Assert.True(reopened.RemoveCustomShow(saved));
                Assert.Empty(reopened.CustomShows);
                reopened.Save();

                using PowerPointPresentation removed =
                    PowerPointPresentation.Load(path);
                Assert.Empty(removed.CustomShows);
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void CustomShows_RejectForeignSlidesShowsDuplicatesAndEmptySequences() {
            using PowerPointPresentation first = PowerPointPresentation.Create();
            PowerPointSlide local = first.AddSlide();
            using PowerPointPresentation second = PowerPointPresentation.Create();
            PowerPointSlide foreign = second.AddSlide();

            PowerPointCustomShow show = first.AddCustomShow("Local",
                new[] { local });
            PowerPointCustomShow foreignShow = second.AddCustomShow("Foreign",
                new[] { foreign });

            Assert.Throws<InvalidOperationException>(() =>
                first.AddCustomShow("Foreign slide", new[] { foreign }));
            Assert.Throws<InvalidOperationException>(() =>
                first.AddCustomShow("LOCAL", new[] { local }));
            Assert.Throws<ArgumentException>(() =>
                first.AddCustomShow("Empty", Array.Empty<PowerPointSlide>()));
            Assert.Throws<InvalidOperationException>(() =>
                show.SetSlides(new[] { foreign }));
            Assert.Throws<InvalidOperationException>(() =>
                first.RenameCustomShow(foreignShow, "Still foreign"));
            Assert.Throws<InvalidOperationException>(() =>
                first.RemoveCustomShow(foreignShow));

            Assert.Throws<ArgumentException>(() =>
                first.AddCustomShow("Invalid\u0001name", new[] { local }));
            Assert.Throws<ArgumentException>(() =>
                first.RenameCustomShow(show, "Invalid\u0001name"));
            Assert.Equal("Local", show.Name);
        }

        [Fact]
        public void FeatureReportPreservesMalformedOrExtendedCustomShows() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointCustomShow show = presentation.AddCustomShow(
                "Primary", new[] { slide });
            CustomShowList list = presentation.OpenXmlDocument.PresentationPart!
                .Presentation!.CustomShowList!;
            var duplicate = new CustomShow {
                Id = show.Id,
                Name = "Duplicate",
                SlideList = new SlideList(new SlideListEntry {
                    Id = "rUnresolved"
                })
            };
            duplicate.Append(new SlideList(new SlideListEntry {
                Id = presentation.OpenXmlDocument.PresentationPart!
                    .GetIdOfPart(slide.SlidePart)
            }));
            list.Append(duplicate);
            list.Append(new OpenXmlUnknownElement("p14", "extLst",
                "http://schemas.microsoft.com/office/powerpoint/2010/main"));

            PowerPointFeatureReport report = presentation.InspectFeatures();
            PowerPointFeatureFinding finding = Assert.Single(
                report.FindFeatures("Custom shows"));

            Assert.Equal(OfficeFeatureSupportLevel.Preserved,
                finding.SupportLevel);
            Assert.Contains(finding.Details, detail => detail.Contains(
                "duplicates identifier", StringComparison.OrdinalIgnoreCase));
            Assert.Contains(finding.Details, detail => detail.Contains(
                "unresolved slide relationship", StringComparison.OrdinalIgnoreCase));
            Assert.Contains(finding.Details, detail => detail.Contains(
                "exactly one", StringComparison.OrdinalIgnoreCase));
            Assert.Contains(finding.Details, detail => detail.Contains(
                "extension", StringComparison.OrdinalIgnoreCase));
            Assert.Throws<InvalidOperationException>(() =>
                report.EnsureNoAdvancedFeatures());
        }

        [Fact]
        public void FeatureReportRetainsExtensionOnlyCustomShowList() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            presentation.AddSlide();
            presentation.OpenXmlDocument.PresentationPart!.Presentation!
                .CustomShowList = new CustomShowList(new OpenXmlUnknownElement(
                    "p14", "extLst",
                    "http://schemas.microsoft.com/office/powerpoint/2010/main"));

            PowerPointFeatureReport report = presentation.InspectFeatures();
            PowerPointFeatureFinding finding = Assert.Single(
                report.FindFeatures("Custom shows"));

            Assert.Equal(OfficeFeatureSupportLevel.Preserved,
                finding.SupportLevel);
            Assert.Equal(1, finding.Count);
            Assert.Throws<InvalidOperationException>(() =>
                report.EnsureNoAdvancedFeatures());
        }

        [Fact]
        public void FeatureReportPreservesCustomShowExtendedAttributes() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            presentation.AddCustomShow("Primary", new[] { slide });
            CustomShowList list = presentation.OpenXmlDocument
                .PresentationPart!.Presentation!.CustomShowList!;
            CustomShow show = list.Elements<CustomShow>().Single();
            SlideList slideList = show.SlideList!;
            SlideListEntry entry = slideList.Elements<SlideListEntry>()
                .Single();
            foreach (OpenXmlElement element in new OpenXmlElement[] {
                         list, show, slideList, entry
                     }) {
                element.SetAttribute(new OpenXmlAttribute("producer",
                    "metadata", "urn:officeimo:test", "retained"));
            }

            PowerPointFeatureFinding finding = Assert.Single(
                presentation.InspectFeatures().FindFeatures("Custom shows"));

            Assert.Equal(OfficeFeatureSupportLevel.Preserved,
                finding.SupportLevel);
            Assert.Contains(finding.Details, detail => detail.Contains(
                "attributes", StringComparison.OrdinalIgnoreCase));
        }

        [Fact]
        public void CustomShowMutationsPreserveProducerListAndEntryMetadata() {
            const string ProducerNamespace = "urn:officeimo:test";
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide first = presentation.AddSlide();
            PowerPointSlide second = presentation.AddSlide();
            PowerPointSlide third = presentation.AddSlide();
            PowerPointCustomShow show = presentation.AddCustomShow(
                "Producer show", new[] { first, second });
            SlideList slideList = show.OpenXmlElement.SlideList!;
            slideList.SetAttribute(new OpenXmlAttribute("producer",
                "metadata", ProducerNamespace, "list"));
            slideList.Append(new OpenXmlUnknownElement("producer",
                "listExtension", ProducerNamespace));
            SlideListEntry firstEntry = slideList
                .Elements<SlideListEntry>().First();
            firstEntry.SetAttribute(new OpenXmlAttribute("producer",
                "metadata", ProducerNamespace, "first"));

            show.SetSlides(new[] { second, first });
            show.AddSlide(third);
            show.InsertSlide(1, first);
            show.MoveSlide(3, 0);
            Assert.True(show.RemoveSlide(second));

            Assert.Equal(new[] { third, first, first }, show.Slides);
            Assert.Equal("list", slideList.GetAttribute("metadata",
                ProducerNamespace).Value);
            Assert.Contains(slideList.ChildElements, child =>
                child.LocalName == "listExtension"
                && child.NamespaceUri == ProducerNamespace);
            string firstRelationshipId = presentation.OpenXmlDocument
                .PresentationPart!.GetIdOfPart(first.SlidePart);
            SlideListEntry preservedFirst = Assert.Single(slideList
                .Elements<SlideListEntry>(), entry =>
                entry.Id?.Value == firstRelationshipId
                && entry.GetAttributes().Any(attribute =>
                    attribute.LocalName == "metadata"
                    && attribute.NamespaceUri == ProducerNamespace
                    && attribute.Value == "first"));
            Assert.Same(firstEntry, preservedFirst);
        }

        [Fact]
        public void CustomShowDuplicateMutationsPreservePositionalEntryMetadata() {
            const string ProducerNamespace = "urn:officeimo:test";
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointCustomShow show = presentation.AddCustomShow(
                "Duplicate show", new[] { slide, slide });
            SlideListEntry[] entries = show.OpenXmlElement.SlideList!
                .Elements<SlideListEntry>().ToArray();
            entries[0].SetAttribute(new OpenXmlAttribute("producer",
                "metadata", ProducerNamespace, "first"));
            entries[1].SetAttribute(new OpenXmlAttribute("producer",
                "metadata", ProducerNamespace, "second"));

            Assert.True(show.RemoveSlide(slide));
            Assert.Same(entries[1], Assert.Single(show.OpenXmlElement
                .SlideList!.Elements<SlideListEntry>()));
            show.InsertSlide(0, slide);
            show.MoveSlide(1, 0);

            SlideListEntry[] finalEntries = show.OpenXmlElement.SlideList!
                .Elements<SlideListEntry>().ToArray();
            Assert.Equal(2, finalEntries.Length);
            Assert.Same(entries[1], finalEntries[0]);
            Assert.Contains(finalEntries[0].GetAttributes(), attribute =>
                attribute.LocalName == "metadata"
                && attribute.NamespaceUri == ProducerNamespace
                && attribute.Value == "second");
            Assert.DoesNotContain(finalEntries[1].GetAttributes(),
                attribute => attribute.LocalName == "metadata"
                    && attribute.NamespaceUri == ProducerNamespace);
        }

        [Fact]
        public void CustomShows_RemoveZeroIdentifierAlsoRemovesTargetingActions() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape actionShape = slide.AddRectanglePoints(
                20, 20, 120, 60);
            PowerPointCustomShow show = presentation.AddCustomShow(
                "Imported zero", new[] { slide });
            show.OpenXmlElement.Id = 0U;
            NonVisualDrawingProperties actionProperties =
                ((Shape)actionShape.Element).NonVisualShapeProperties!
                .NonVisualDrawingProperties!;
            actionProperties.Append(new A.HyperlinkOnClick {
                Id = string.Empty,
                Action = "ppaction://customshow?id=0&return=true"
            });

            Assert.True(presentation.RemoveCustomShow(show));

            Assert.Empty(slide.SlidePart.Slide!
                .Descendants<A.HyperlinkOnClick>());
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void CustomShows_ReuseFreeIdentifierWhenMaximumIsOccupied() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointCustomShow maximum = presentation.AddCustomShow(
                "Maximum", new[] { slide });
            maximum.OpenXmlElement.Id = uint.MaxValue;

            PowerPointCustomShow allocated = presentation.AddCustomShow(
                "Available", new[] { slide });

            Assert.Equal(1U, allocated.Id);
            Assert.Equal(new[] { uint.MaxValue, 1U },
                presentation.CustomShows.Select(show => show.Id));
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void CustomShows_RemoveActionRelationshipsOwnedByDeletedShow() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide source = presentation.AddSlide();
            PowerPointSlide target = presentation.AddSlide();
            PowerPointCustomShow show = presentation.AddCustomShow(
                "Linked", new[] { target });
            HyperlinkRelationship external = source.SlidePart
                .AddHyperlinkRelationship(
                    new Uri("https://example.test/custom-show"), true);
            const string InternalRelationshipId = "rIdCustomShowTarget";
            source.SlidePart.AddPart(target.SlidePart,
                InternalRelationshipId);
            AppendCustomShowAction(source.AddRectanglePoints(
                    20, 20, 120, 60), show.Id, external.Id);
            AppendCustomShowAction(source.AddRectanglePoints(
                    20, 100, 120, 60), show.Id,
                InternalRelationshipId);

            Assert.True(presentation.RemoveCustomShow(show));

            Assert.Empty(source.SlidePart.Slide!
                .Descendants<A.HyperlinkOnClick>());
            Assert.Empty(source.SlidePart.HyperlinkRelationships);
            Assert.DoesNotContain(source.SlidePart.Parts,
                pair => pair.RelationshipId == InternalRelationshipId);
            Assert.Equal(2, presentation.Slides.Count);
            Assert.Empty(presentation.ValidateDocument());
        }

        private static void AppendCustomShowAction(PowerPointAutoShape shape,
            uint showId, string relationshipId) {
            NonVisualDrawingProperties properties = ((Shape)shape.Element)
                .NonVisualShapeProperties!.NonVisualDrawingProperties!;
            properties.Append(new A.HyperlinkOnClick {
                Id = relationshipId,
                Action = "ppaction://customshow?id=" + showId
            });
        }
    }
}
