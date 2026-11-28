using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests {
    public sealed class PowerPointCustomShowTests {
        [Fact]
        public void OwnedCustomShowsIgnoreExtensionDecoysAndKeepHeldShowsAcrossLateSdkBinding() {
            using var source = PowerPointDocument.Create();
            PowerPointSlide first = source.AddSlide();
            PowerPointSlide second = source.AddSlide();
            source.AddCustomShow("Visible", new[] { first, second });
            byte[] bytes = source.ToBytes();
            using var read = OfficeOpcPackage.Open(new MemoryStream(bytes, writable: false));
            string main = read.GetRelatedPartName(null, OfficeOpcPackage.OfficeDocumentRelationshipType)!;
            using var edit = OfficeOpcEditablePackage.Open(new MemoryStream(bytes, writable: false));
            OfficeXmlTree tree = edit.GetXml(main);
            const int ns = OfficeXmlNamespaces.PresentationML;
            int extensions = tree.FindChild(tree.Root, ns, "extLst");
            if (extensions == OfficeXmlTree.NoNode) {
                extensions = tree.CreateElement(ns, "extLst");
                tree.Append(tree.Root, extensions);
            }
            int extension = tree.CreateElement(ns, "ext");
            tree.SetAttribute(extension, OfficeXmlNamespaces.None, "uri", "urn:fixture:unknown-custom-show");
            tree.Append(extensions, extension);
            int shadow = tree.CreateElement(ns, "custShowLst");
            tree.Append(extension, shadow);
            int decoy = tree.CreateElement(ns, "custShow");
            tree.SetAttribute(decoy, OfficeXmlNamespaces.None, "name", "Shadow");
            tree.SetAttribute(decoy, OfficeXmlNamespaces.None, "id", "2");
            tree.Append(shadow, decoy);
            using var packaged = new MemoryStream();
            edit.Save(packaged);

            using var document = PowerPointDocument.Load(new MemoryStream(packaged.ToArray(), writable: false));
            Assert.False(document.IsSdkLoaded, "load");
            PowerPointCustomShow held = Assert.Single(document.CustomShows);
            Assert.Null(document.GetCustomShow("Shadow"));
            Assert.False(document.IsSdkLoaded, "read");
            PowerPointCustomShow newShow = document.AddCustomShow("New show", new[] { document.Slides[1] });
            Assert.Equal(2U, newShow.Id);
            Assert.False(document.IsSdkLoaded, "create");
            held.MoveSlide(0, 1);
            document.RenameCustomShow(held, "Renamed");
            Assert.Equal(new[] { document.Slides[1], document.Slides[0] }, held.Slides);
            Assert.False(document.IsSdkLoaded, "edit");

            byte[] saved = document.ToBytes();
            Assert.False(document.IsSdkLoaded, "save");
            using var reopened = PowerPointDocument.Load(new MemoryStream(saved, writable: false));
            Assert.Equal(new[] { "Renamed", "New show" }, reopened.CustomShows.Select(show => show.Name));
            Assert.Null(reopened.GetCustomShow("Shadow"));
            Assert.False(reopened.IsSdkLoaded);
            using var savedPackage = OfficeOpcEditablePackage.Open(new MemoryStream(saved, writable: false));
            OfficeXmlTree savedPresentation = savedPackage.GetXml(main);
            int savedExtensions = savedPresentation.FindChild(savedPresentation.Root, ns, "extLst");
            int savedExtension = savedPresentation.FindChild(savedExtensions, ns, "ext");
            int savedShadow = savedPresentation.FindChild(savedExtension, ns, "custShowLst");
            int savedDecoy = savedPresentation.FindChild(savedShadow, ns, "custShow");
            Assert.Equal("Shadow", savedPresentation.GetAttribute(savedDecoy, OfficeXmlNamespaces.None, "name"));

            _ = document.OpenXmlDocument;
            Assert.Equal("Renamed", held.Name);
            Assert.True(document.RemoveCustomShow(held));
            Assert.Equal("Renamed", held.Name);
            Assert.Equal("New show", Assert.Single(document.CustomShows).Name);
        }

        [Fact]
        public void CustomShowSequenceFollowsAnSdkBindDuringSlideEnumeration() {
            using var source = PowerPointDocument.Create();
            PowerPointSlide first = source.AddSlide();
            source.AddSlide();
            source.AddCustomShow("Sequence", new[] { first });
            using var document = PowerPointDocument.Load(new MemoryStream(source.ToBytes(), writable: false));
            PowerPointCustomShow show = Assert.Single(document.CustomShows);
            Assert.False(document.IsSdkLoaded);

            show.SetSlides(SelectSlides());
            Assert.Equal(new[] { document.Slides[1] }, show.Slides);
            using var reopened = PowerPointDocument.Load(new MemoryStream(document.ToBytes(), writable: false));
            Assert.Equal(new[] { reopened.Slides[1] }, Assert.Single(reopened.CustomShows).Slides);

            IEnumerable<PowerPointSlide> SelectSlides() {
                _ = document.OpenXmlDocument;
                yield return document.Slides[1];
            }
        }

        [Fact]
        public void SdkCustomShowEditsRetainUnknownEntriesAndAdjacentComments() {
            using var document = PowerPointDocument.Create();
            PowerPointSlide first = document.AddSlide();
            PowerPointSlide second = document.AddSlide();
            PowerPointCustomShow show = document.AddCustomShow("Sequence", new[] { first });
            SlideList list = GetSdkShow(document, show).SlideList!;
            SlideListEntry typed = Assert.Single(list.Elements<SlideListEntry>());
            var unknown = new OpenXmlUnknownElement("p", "sld", typed.NamespaceUri);
            unknown.SetAttribute(new OpenXmlAttribute("fixture", "metadata", "urn:fixture", "untyped"));
            var comment = new OpenXmlMiscNode(System.Xml.XmlNodeType.Comment, "<!--producer note-->");
            list.Append(comment, unknown);

            show.AddSlide(second);
            Assert.IsType<SlideListEntry>(list.ChildElements[1]);
            Assert.Same(comment, list.ChildElements[2]);
            Assert.Same(unknown, list.ChildElements[3]);

            show.SetSlides(new[] { first });
            Assert.Same(typed, list.ChildElements[0]);
            Assert.Same(comment, list.ChildElements[1]);
            Assert.Same(unknown, list.ChildElements[2]);
            Assert.Equal(new[] { first }, show.Slides);
        }

        [Fact]
        public void CustomShows_CreateEditAndRoundTripThroughPublicApi() {
            string path = Path.Combine(Path.GetTempPath(),
                "OfficeIMO-CustomShows-" + Guid.NewGuid().ToString("N") + ".pptx");
            try {
                using (PowerPointDocument presentation =
                       PowerPointDocument.Create(path)) {
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

                using PowerPointDocument reopened =
                    PowerPointDocument.Load(path);
                PowerPointCustomShow saved = Assert.Single(reopened.CustomShows);
                Assert.Equal("Decision path", saved.Name);
                Assert.Equal(new[] { reopened.Slides[0], reopened.Slides[2] },
                    saved.Slides);
                Assert.Equal(saved.Id,
                    reopened.GetCustomShow("decision PATH")!.Id);

                Assert.True(reopened.RemoveCustomShow(saved));
                Assert.Empty(reopened.CustomShows);
                reopened.Save();

                using PowerPointDocument removed =
                    PowerPointDocument.Load(path);
                Assert.Empty(removed.CustomShows);
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void CustomShows_RejectForeignSlidesShowsDuplicatesAndEmptySequences() {
            using PowerPointDocument first = PowerPointDocument.Create();
            PowerPointSlide local = first.AddSlide();
            using PowerPointDocument second = PowerPointDocument.Create();
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
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
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
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
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
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
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
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            PowerPointSlide first = presentation.AddSlide();
            PowerPointSlide second = presentation.AddSlide();
            PowerPointSlide third = presentation.AddSlide();
            PowerPointCustomShow show = presentation.AddCustomShow(
                "Producer show", new[] { first, second });
            SlideList slideList = GetSdkShow(presentation, show).SlideList!;
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
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointCustomShow show = presentation.AddCustomShow(
                "Duplicate show", new[] { slide, slide });
            SlideList slideList = GetSdkShow(presentation, show).SlideList!;
            SlideListEntry[] entries = slideList.Elements<SlideListEntry>().ToArray();
            entries[0].SetAttribute(new OpenXmlAttribute("producer",
                "metadata", ProducerNamespace, "first"));
            entries[1].SetAttribute(new OpenXmlAttribute("producer",
                "metadata", ProducerNamespace, "second"));

            Assert.True(show.RemoveSlide(slide));
            Assert.Same(entries[1], Assert.Single(slideList.Elements<SlideListEntry>()));
            show.InsertSlide(0, slide);
            show.MoveSlide(1, 0);

            SlideListEntry[] finalEntries = slideList
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
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape actionShape = slide.AddRectangle(
                OfficeIMO.OfficeLength.Points(20), OfficeIMO.OfficeLength.Points(20), OfficeIMO.OfficeLength.Points(120), OfficeIMO.OfficeLength.Points(60));
            PowerPointCustomShow show = presentation.AddCustomShow(
                "Imported zero", new[] { slide });
            GetSdkShow(presentation, show).Id = 0U;
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
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointCustomShow maximum = presentation.AddCustomShow(
                "Maximum", new[] { slide });
            GetSdkShow(presentation, maximum).Id = uint.MaxValue;

            PowerPointCustomShow allocated = presentation.AddCustomShow(
                "Available", new[] { slide });

            Assert.Equal(1U, allocated.Id);
            Assert.Equal(new[] { uint.MaxValue, 1U },
                presentation.CustomShows.Select(show => show.Id));
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void CustomShows_RemoveActionRelationshipsOwnedByDeletedShow() {
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
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
            AppendCustomShowAction(source.AddRectangle(
                    OfficeIMO.OfficeLength.Points(20), OfficeIMO.OfficeLength.Points(20), OfficeIMO.OfficeLength.Points(120), OfficeIMO.OfficeLength.Points(60)), show.Id, external.Id);
            AppendCustomShowAction(source.AddRectangle(
                    OfficeIMO.OfficeLength.Points(20), OfficeIMO.OfficeLength.Points(100), OfficeIMO.OfficeLength.Points(120), OfficeIMO.OfficeLength.Points(60)), show.Id,
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

        [Fact]
        public void OwnedCustomShowRemovalKeepsOtherActionsAndExtensionMarkup() {
            using var source = PowerPointDocument.Create();
            PowerPointSlide first = source.AddSlide();
            PowerPointSlide second = source.AddSlide();
            PowerPointCustomShow removed = source.AddCustomShow("Remove", new[] { second });
            PowerPointCustomShow kept = source.AddCustomShow("Keep", new[] { first });
            const string alternateId = "rIdCustomShowAlternate";
            const string keepId = "rIdCustomShowKeep";
            const string externalId = "rIdCustomShowExternal";
            const string internalId = "rIdCustomShowInternal";
            first.SlidePart.AddHyperlinkRelationship(new Uri("https://example.test/remove"), true, externalId);
            first.SlidePart.AddHyperlinkRelationship(new Uri("https://example.test/keep"), true, keepId);
            first.SlidePart.AddHyperlinkRelationship(new Uri("https://example.test/alternate"), true, alternateId);
            first.SlidePart.AddPart(second.SlidePart, internalId);
            AppendCustomShowAction(first.AddRectangle(OfficeLength.Points(20), OfficeLength.Points(20), OfficeLength.Points(120), OfficeLength.Points(40)), removed.Id, externalId);
            AppendCustomShowAction(first.AddRectangle(OfficeLength.Points(20), OfficeLength.Points(80), OfficeLength.Points(120), OfficeLength.Points(40)), removed.Id, internalId);
            AppendCustomShowAction(first.AddRectangle(OfficeLength.Points(20), OfficeLength.Points(140), OfficeLength.Points(120), OfficeLength.Points(40)), kept.Id, keepId);
            var text = (Shape)first.AddTextBox("Run action").Element;
            A.Run run = Assert.Single(text.TextBody!.Descendants<A.Run>());
            run.RunProperties ??= new A.RunProperties();
            run.RunProperties.Append(new A.HyperlinkOnMouseOver {
                Id = externalId,
                Action = "ppaction://customshow?id=" + removed.Id
            });

            string slidePart = first.SlidePart.Uri.OriginalString;
            using var input = OfficeOpcEditablePackage.Open(new MemoryStream(source.ToBytes(), writable: false));
            OfficeXmlTree tree = input.GetXml(slidePart);
            const int p = OfficeXmlNamespaces.PresentationML, a = OfficeXmlNamespaces.DrawingML;
            const int mc = OfficeXmlNamespaces.MarkupCompatibility, r = OfficeXmlNamespaces.Relationships;
            int extensions = tree.FindChild(tree.Root, p, "extLst");
            if (extensions == OfficeXmlTree.NoNode) extensions = Append(tree.Root, p, "extLst");
            int extension = Append(extensions, p, "ext");
            tree.SetAttribute(extension, OfficeXmlNamespaces.None, "uri", "urn:fixture:action-decoy");
            SetDecoy(Append(extension, a, "hlinkClick"), alternateId);
            int alternate = Append(tree.Root, mc, "AlternateContent");
            int choice = Append(alternate, mc, "Choice");
            tree.SetAttribute(choice, OfficeXmlNamespaces.Xmlns, "future", "urn:fixture:unsupported");
            tree.SetAttribute(choice, OfficeXmlNamespaces.None, "Requires", "future");
            SetDecoy(Append(choice, a, "hlinkClick"), alternateId);
            Append(alternate, mc, "Fallback");
            using var packaged = new MemoryStream();
            input.Save(packaged);

            using var presentation = PowerPointDocument.Load(new MemoryStream(packaged.ToArray(), writable: false));
            PowerPointCustomShow held = presentation.GetCustomShow("Remove")!;
            PowerPointCustomShow secondHandle = presentation.GetCustomShow("Remove")!;
            Assert.True(presentation.RemoveCustomShow(held));
            Assert.Equal("Keep", Assert.Single(presentation.CustomShows).Name);
            Assert.False(presentation.IsSdkLoaded);

            byte[] bytes = presentation.ToBytes();
            Assert.False(presentation.IsSdkLoaded);
            using var persisted = OfficeOpcEditablePackage.Open(new MemoryStream(bytes, writable: false));
            OfficeXmlTree savedTree = persisted.GetXml(slidePart);
            OfficeXmlNode savedRoot = new OfficeXmlNode(savedTree, savedTree.Root);
            OfficeXmlNode action = Assert.Single(savedRoot.ContentDescendants(a, "hlinkClick"));
            Assert.Equal(keepId, action.Attribute(r, "id"));
            Assert.Empty(savedRoot.ContentDescendants(a, "hlinkMouseOver"));
            Assert.Equal(2, savedRoot.Descendants(a, "hlinkClick").Count(link => link.Attribute(r, "id") == alternateId));
            OfficeXmlTree relationships = persisted.GetRelationships(slidePart);
            int ns = OfficeXmlNamespaces.PackageRelationships;
            var ids = new List<string>();
            for (int edge = relationships.FindChild(relationships.Root, ns, "Relationship"); edge != OfficeXmlTree.NoNode;
                edge = relationships.FindNextSibling(edge, ns, "Relationship")) {
                ids.Add(relationships.GetAttribute(edge, OfficeXmlNamespaces.None, "Id")!);
            }
            Assert.Contains(keepId, ids);
            Assert.Contains(alternateId, ids);
            Assert.DoesNotContain(externalId, ids);
            Assert.DoesNotContain(internalId, ids);

            _ = presentation.OpenXmlDocument;
            Assert.Equal("Remove", secondHandle.Name);
            Assert.Equal(new[] { presentation.Slides[1] }, secondHandle.Slides);

            int Append(int parent, int namespaceId, string name) {
                int child = tree.CreateElement(namespaceId, name);
                tree.Append(parent, child);
                return child;
            }

            void SetDecoy(int node, string relationship) {
                tree.SetAttribute(node, OfficeXmlNamespaces.None, "action", "ppaction://customshow?id=" + removed.Id);
                tree.SetAttribute(node, r, "id", relationship);
            }
        }

        [Fact]
        public void CustomShowRemovalKeepsSdkCleanupForTypedAlternateShapes() {
            using var source = PowerPointDocument.Create();
            PowerPointSlide slide = source.AddSlide();
            PowerPointCustomShow show = source.AddCustomShow("Alternates", new[] { slide });
            PowerPointAutoShape rectangle = slide.AddRectangle(OfficeLength.Points(20), OfficeLength.Points(20),
                OfficeLength.Points(120), OfficeLength.Points(40));
            const string relationshipId = "rIdAlternateAction";
            slide.SlidePart.AddHyperlinkRelationship(new Uri("https://example.test/alternate"), true, relationshipId);
            AppendCustomShowAction(rectangle, show.Id, relationshipId);

            var shape = (Shape)rectangle.Element;
            var shapeTree = (ShapeTree)shape.Parent!;
            shape.Remove();
            var choice = new AlternateContentChoice { Requires = "future" };
            choice.AddNamespaceDeclaration("future", "urn:fixture:unsupported");
            shapeTree.Append(new AlternateContent(choice, new AlternateContentFallback(shape)));

            byte[] bytes = source.ToBytes();
            using var sdk = PowerPointDocument.Load(new MemoryStream(bytes, writable: false));
            SlidePart sdkSlide = sdk.Slides[0].SlidePart;
            Assert.Single(sdkSlide.Slide!.Descendants<A.HyperlinkOnClick>());
            Assert.True(sdk.RemoveCustomShow(Assert.Single(sdk.CustomShows)));
            Assert.DoesNotContain(sdkSlide.HyperlinkRelationships, item => item.Id == relationshipId);

            using var presentation = PowerPointDocument.Load(new MemoryStream(bytes, writable: false));
            Assert.False(presentation.IsSdkLoaded);
            Assert.True(presentation.RemoveCustomShow(Assert.Single(presentation.CustomShows)));
            Assert.True(presentation.IsSdkLoaded);
            Assert.Empty(presentation.Slides[0].SlidePart.Slide!.Descendants<A.HyperlinkOnClick>());
            Assert.DoesNotContain(presentation.Slides[0].SlidePart.HyperlinkRelationships, item => item.Id == relationshipId);
        }

        [Fact]
        public void CustomShowRemovalPreservesAnOpaquePartWithTheSlideContentType() {
            using var source = PowerPointDocument.Create();
            PowerPointSlide slide = source.AddSlide();
            PowerPointCustomShow show = source.AddCustomShow("Opaque", new[] { slide });
            source.AddCustomShow("Keep", new[] { slide });
            PowerPointAutoShape rectangle = slide.AddRectangle(OfficeLength.Points(20), OfficeLength.Points(20),
                OfficeLength.Points(120), OfficeLength.Points(40));
            AppendCustomShowAction(rectangle, show.Id, string.Empty);
            SlidePart actualSlide = slide.SlidePart;
            var encoding = System.Text.Encoding.Unicode;
            byte[] opaqueBytes = encoding.GetPreamble().Concat(encoding.GetBytes(
                "<?xml version=\"1.0\" encoding=\"utf-16\"?>" + actualSlide.Slide!.OuterXml)).ToArray();
            ExtendedPart opaque = source.OpenXmlDocument.PresentationPart!.AddExtendedPart(
                "urn:fixture:custom-slide", actualSlide.ContentType, ".xml");
            using (var input = new MemoryStream(opaqueBytes, writable: false)) opaque.FeedData(input);
            string opaqueName = opaque.Uri.OriginalString;

            using var presentation = PowerPointDocument.Load(new MemoryStream(source.ToBytes(), writable: false));
            presentation.RenameCustomShow(presentation.GetCustomShow("Keep")!, "Pending rename");
            Assert.False(presentation.IsSdkLoaded);
            Assert.True(presentation.RemoveCustomShow(presentation.GetCustomShow("Opaque")!));
            Assert.Equal("Pending rename", Assert.Single(presentation.CustomShows).Name);
            byte[] bytes = presentation.ToBytes();
            using var saved = OfficeOpcEditablePackage.Open(new MemoryStream(bytes, writable: false));
            Assert.Equal(opaqueBytes, saved.ReadPartBytes(opaqueName));
            OfficeXmlTree actual = saved.GetXml(actualSlide.Uri.OriginalString);
            Assert.Empty(new OfficeXmlNode(actual, actual.Root).ContentDescendants(OfficeXmlNamespaces.DrawingML, "hlinkClick"));
        }

        private static CustomShow GetSdkShow(PowerPointDocument presentation, PowerPointCustomShow show) {
            uint id = show.Id;
            return Assert.Single(presentation.OpenXmlDocument.PresentationPart!.Presentation!
                .CustomShowList!.Elements<CustomShow>(), element => element.Id?.Value == id);
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
