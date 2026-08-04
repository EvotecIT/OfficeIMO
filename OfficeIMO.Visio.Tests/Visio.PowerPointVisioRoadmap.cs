using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;
using OfficeIMO.Visio.Pdf;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioPowerPointVisioRoadmapTests {
        [Theory]
        [InlineData(VisioPackageType.Drawing, ".vsdx")]
        [InlineData(VisioPackageType.Template, ".vstx")]
        [InlineData(VisioPackageType.Stencil, ".vssx")]
        [InlineData(VisioPackageType.MacroEnabledDrawing, ".vsdm")]
        [InlineData(VisioPackageType.MacroEnabledTemplate, ".vstm")]
        [InlineData(VisioPackageType.MacroEnabledStencil, ".vssm")]
        public void OpenXmlPackageFamiliesLoadRoundTripAndConvert(
            VisioPackageType type, string extension) {
            string source = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + extension);
            string roundTrip = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + extension);
            string pdf = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pdf");
            try {
                VisioDocument authored = VisioDocument.Create(source, type);
                authored.AddPage("Representative", 8.5, 6)
                    .AddRectangle(3, 3, 2, 1, type.ToString());
                authored.Save();
                VisioDocument loaded = VisioDocument.Load(source);
                Assert.Equal(type, loaded.PackageType);
                Assert.Equal(type.ToString(), loaded.Pages.Single().Shapes.Single().Text);
                Assert.Single(loaded.ToOfficeDocumentReadResult().Pages);
                OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
                    .AddVisioHandler().Build();
                Assert.Single(reader.ReadDocument(source).Pages);
                loaded.Save(roundTrip, type);
                Assert.Equal(type, VisioDocument.Load(roundTrip).PackageType);
                loaded.SaveAsPdf(pdf);
                Assert.StartsWith("%PDF-", Encoding.ASCII.GetString(File.ReadAllBytes(pdf), 0, 5));
            } finally {
                if (File.Exists(source)) File.Delete(source);
                if (File.Exists(roundTrip)) File.Delete(roundTrip);
                if (File.Exists(pdf)) File.Delete(pdf);
            }
        }

        [Fact]
        public void MacroPayloadAndRelationshipSubtreeAreOpaqueAndPreservedAcrossLoadSave() {
            string source = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdm");
            string saved = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdm");
            byte[] payload = { 0xD0, 0xCF, 0x11, 0xE0, 1, 2, 3, 4, 5 };
            try {
                VisioDocument document = VisioDocument.Create(source,
                    VisioPackageType.MacroEnabledDrawing);
                document.AddPage("Macro", 8.5, 6).AddRectangle(2, 2, 1, 1, "Opaque VBA");
                document.Save();
                InjectVbaProject(source, payload);
                VisioDocument loaded = VisioDocument.Load(source);
                Assert.True(loaded.HasVbaProject);
                Assert.True(loaded.IsMacroEnabled);
                loaded.Save(saved);
                using ZipArchive archive = ZipFile.OpenRead(saved);
                Assert.Null(archive.GetEntry("visio/vbaProject.bin"));
                using Stream input = archive.GetEntry(
                    "visio/macros/project.bin")!.Open();
                using var copy = new MemoryStream();
                input.CopyTo(copy);
                Assert.Equal(payload, copy.ToArray());
                ZipArchiveEntry signature = archive.GetEntry(
                    "visio/macros/signature.bin")!;
                using Stream signatureInput = signature.Open();
                using var signatureCopy = new MemoryStream();
                signatureInput.CopyTo(signatureCopy);
                Assert.Equal(new byte[] { 9, 8, 7, 6 },
                    signatureCopy.ToArray());
                string relationships;
                using (var reader = new StreamReader(archive.GetEntry(
                           "visio/macros/_rels/project.bin.rels")!.Open()))
                    relationships = reader.ReadToEnd();
                Assert.Contains("vbaProjectSignature", relationships,
                    StringComparison.Ordinal);
                Assert.Contains("https://example.test/signing-policy",
                    relationships, StringComparison.Ordinal);
                using (var reader = new StreamReader(archive.GetEntry(
                           "[Content_Types].xml")!.Open()))
                    relationships = reader.ReadToEnd();
                Assert.Contains("/visio/macros/project.bin", relationships,
                    StringComparison.Ordinal);
                Assert.DoesNotContain("/visio/vbaProject.bin", relationships,
                    StringComparison.Ordinal);
            } finally {
                if (File.Exists(source)) File.Delete(source);
                if (File.Exists(saved)) File.Delete(saved);
            }
        }

        [Fact]
        public void PageLessStencilWithMastersLoadsAndReadsWithoutInventingPages() {
            string source = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vssx");
            try {
                VisioDocument authored = VisioDocument.Create(source, VisioPackageType.Stencil);
                authored.RegisterMaster("Stencil Node", new VisioShape("1", 1, 1, 1, 1,
                    "Reusable"));
                authored.AddPage("Temporary", 8.5, 6).AddRectangle(2, 2, 1, 1, "Master seed");
                authored.Save();
                RemovePagesFromStencil(source);

                VisioDocument loaded = VisioDocument.Load(source);
                Assert.Equal(VisioPackageType.Stencil, loaded.PackageType);
                Assert.Empty(loaded.Pages);
                Assert.Empty(loaded.ToOfficeDocumentReadResult().Pages);
                loaded.Save();
                VisioDocument roundTripped = VisioDocument.Load(source);
                Assert.Empty(roundTripped.Pages);
                Assert.NotEmpty(roundTripped.Masters);
            } finally {
                if (File.Exists(source)) File.Delete(source);
            }
        }

        [Fact]
        public void ShapeSheetDataGraphicsNestedContainersAndThreadsRoundTrip() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            try {
                VisioDocument document = VisioDocument.Create(path);
                VisioPage page = document.AddPage("Typed", 11, 8.5);
                VisioShape target = page.AddRectangle(3, 5, 1.5, 0.8, "API");
                target.SetShapeData("Status", "Healthy");
                target.SetShapeData("Slo", "91");
                var section = new VisioShapeSheetSection("Actions");
                section.GetOrAddRow("Refresh").SetCell("Action",
                    formula: "CALLTHIS(\"Refresh\")");
                target.SetShapeSheetSection(section);

                VisioDataGraphic definition = VisioDataGraphic.Create()
                    .Badge("Status").Bar("Slo", maximumValue: 100, label: "SLO");
                page.AddDataGraphics(target, definition);
                Assert.NotEmpty(page.GetDataGraphic(target).Shapes);
                Assert.Equal(4, page.AddDataGraphicLegend("legend", "Health", definition, 8, 6).Shapes.Count);

                VisioShape inner = page.AddContainer("inner", "Inner", new[] { target });
                VisioShape outer = page.AddContainer("outer", "Outer", new[] { inner });
                page.AddNestedContainer(outer, inner);
                VisioContainerInfo info = page.GetContainerInfo(inner);
                Assert.Equal(1, info.NestingDepth);
                Assert.Contains("outer", info.ParentContainerIds);

                VisioComment root = page.AddComment(target, "Review", "Owner", "OW");
                VisioComment reply = page.ReplyToComment(root.Id, "Done",
                    new VisioCommentAuthor("Reviewer", "RV", "reviewer@example.test"));
                Assert.Equal(root.Id, reply.ParentCommentId);
                Assert.Single(page.GetCommentThreads());
                document.Save();

                VisioDocument loaded = VisioDocument.Load(path);
                VisioPage loadedPage = loaded.Pages.Single();
                VisioShape loadedTarget = loadedPage.Shapes.Single(shape => shape.Text == "API");
                VisioShapeSheetCell action = loadedTarget.GetShapeSheetSections()
                    .Single(sectionItem => sectionItem.Name == "Actions")
                    .FindRow("Refresh")!.FindCell("Action")!;
                Assert.Equal("CALLTHIS(\"Refresh\")", action.Formula);
                Assert.Equal(2, loadedPage.GetCommentThreads().Single().Comments.Count);
                Assert.Single(loadedPage.CommentsByAuthor(
                    new VisioCommentAuthor("Reviewer", "RV", "reviewer@example.test")));
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void SwimlaneGeometryAssignmentIsDeterministicAndReportsOutsideShapes() {
            VisioDocument document = VisioDocument.Create()
                .SwimlaneDiagram("Flow", swim => swim
                    .Lane("sales", "Sales").Lane("ops", "Operations")
                    .Phase("review", "Review").Phase("fulfill", "Fulfill")
                    .Step("work", "Work", "ops", "fulfill"));
            VisioPage page = document.Pages.Single();
            VisioShape work = page.FindShapeById("work")!;
            work.SetUserCell(VisioSemanticUserCells.SwimlaneLaneId, null);
            work.SetUserCell(VisioSemanticUserCells.SwimlanePhaseId, null);
            VisioShape outside = new("outside", -20, -20, 1, 1, "Outside");
            outside.SetUserCell(VisioSemanticUserCells.Kind,
                VisioSemanticUserCells.SwimlaneActivityKind, "STR");
            page.Shapes.Add(outside);
            VisioSwimlaneAssignmentResult result = page.AssignSwimlaneActivities();
            Assert.Contains(result.Assigned, item => item.Shape.Id == "work" &&
                item.LaneId == "ops" && item.PhaseId == "fulfill");
            Assert.Contains("outside", result.UnassignedShapeIds);
            Assert.False(result.Complete);
        }

        [Fact]
        public void WholeDiagramRelayoutHandlesDenseTopologyWithinLinearBudget() {
            VisioDocument document = VisioDocument.Create();
            VisioPage page = document.AddPage("Dense", 11, 8.5);
            VisioShape? previous = null;
            for (int index = 0; index < 1000; index++) {
                VisioShape shape = new("n" + index, index % 13, index % 17,
                    1.2, 0.5, "Node " + index);
                page.Shapes.Add(shape);
                if (previous != null) page.AddConnector(previous, shape, ConnectorKind.Dynamic);
                previous = shape;
            }
            var timer = Stopwatch.StartNew();
            page.RelayoutDiagram(new VisioWholeDiagramRelayoutOptions {
                PolishAfterLayout = false,
                RouteConnectors = false
            });
            timer.Stop();
            Assert.True(timer.Elapsed < TimeSpan.FromSeconds(5),
                $"Dense relayout took {timer.Elapsed}.");
            Assert.True(page.Shapes.Last().PinX > page.Shapes.First().PinX);
            Assert.Equal(1000, page.Shapes.Select(shape => shape.PinX).Distinct().Count());
        }

        [Fact]
        public void WholeDiagramRelayoutKeepsCycleTogetherAndPlacesTailLater() {
            VisioDocument document = VisioDocument.Create();
            VisioPage page = document.AddPage("Cycle", 11, 8.5);
            VisioShape a = page.AddRectangle(1, 5, 1, 0.5, "A");
            VisioShape b = page.AddRectangle(2, 4, 1, 0.5, "B");
            VisioShape c = page.AddRectangle(3, 3, 1, 0.5, "C");
            page.AddConnector(a, b);
            page.AddConnector(b, a);
            page.AddConnector(b, c);
            page.RelayoutDiagram(new VisioWholeDiagramRelayoutOptions {
                RouteConnectors = false, PolishAfterLayout = false
            });
            Assert.Equal(a.PinX, b.PinX, 5);
            Assert.True(c.PinX > a.PinX);
        }

        [Fact]
        public void FitToContentUsesNestedGroupTransformsInPageCoordinates() {
            VisioDocument document = VisioDocument.Create();
            VisioPage page = document.AddPage("Nested", 20, 20);
            var group = new VisioShape("group", 10, 10, 4, 4, string.Empty) {
                Type = "Group", LocPinX = 2, LocPinY = 2
            };
            var child = new VisioShape("child", 8, 2, 2, 2, "Far child") {
                LocPinX = 1, LocPinY = 1
            };
            group.Children.Add(child);
            page.Shapes.Add(group);

            OfficeIMO.Visio.VisioShapeBounds childPageBounds =
                child.GetPageShapeBounds();
            Assert.Equal(15D, childPageBounds.Left, 5);
            Assert.Equal(17D, childPageBounds.Right, 5);
            page.FitToContent(new VisioFitToContentOptions {
                HorizontalMargin = 0.5D,
                VerticalMargin = 0.5D,
                IncludeGroupChildren = true,
                IncludeConnectors = false,
                MoveContent = false
            });
            Assert.Equal(17.5D, page.Width, 5);
        }

        [Fact]
        public void NestedContainerRefitUsesParentCoordinates() {
            VisioDocument document = VisioDocument.Create();
            VisioPage page = document.AddPage("Nested containers", 20, 20);
            var group = new VisioShape("group", 10, 10, 8, 6, string.Empty) {
                Type = "Group", LocPinX = 4, LocPinY = 3,
                Angle = Math.PI / 6D
            };
            var container = new VisioShape("container", 2, 2, 1, 1,
                "Container");
            container.SetUserCell("msvStructureType", "Container", "STR");
            var member = new VisioShape("member", 5, 4, 1.4, 0.8,
                "Nested member");
            group.Children.Add(container);
            group.Children.Add(member);
            page.Shapes.Add(group);
            page.AddToContainer(container, member, resizeToFit: false);
            page.RefitContainer(container, new VisioContainerOptions {
                Margin = 0.2D,
                HeadingHeight = 0D
            });
            OfficeIMO.Visio.VisioShapeBounds outer =
                container.GetPageShapeBounds();
            OfficeIMO.Visio.VisioShapeBounds inner =
                member.GetPageShapeBounds();
            Assert.True(outer.Left <= inner.Left + 1e-6);
            Assert.True(outer.Right >= inner.Right - 1e-6);
            Assert.True(outer.Bottom <= inner.Bottom + 1e-6);
            Assert.True(outer.Top >= inner.Top - 1e-6);
            Assert.Equal(1.8D, container.Width, 5);
            Assert.Equal(1.2D, container.Height, 5);
        }

        [Fact]
        public void CommentThreadInspectionIsReadOnlyAndRootRemovalCascades() {
            VisioDocument document = VisioDocument.Create();
            VisioPage page = document.AddPage("Comments", 8.5, 6);
            VisioComment unthreaded = page.AddComment("Read only", "A", "AA");
            Assert.Null(unthreaded.ThreadId);
            Assert.Single(page.GetCommentThreads());
            Assert.Null(unthreaded.ThreadId);

            VisioComment root = page.AddComment("Root", "A", "AA");
            page.ReplyToComment(root.Id, "Reply",
                new VisioCommentAuthor("B", "BB"));
            Assert.True(page.RemoveComment(root.Id));
            Assert.DoesNotContain(page.Comments,
                comment => comment.Id == root.Id ||
                           comment.ParentCommentId == root.Id);

            VisioShape target = page.AddRectangle(2, 2, 1, 1, "Target");
            VisioComment targeted = page.AddComment(target, "Target root",
                "A", "AA");
            VisioComment detachedReply = page.ReplyToComment(targeted.Id,
                "Imported reply", new VisioCommentAuthor("B", "BB"));
            detachedReply.ShapeId = null;
            Assert.Equal(2, page.RemoveCommentsForShape(target.Id));
            Assert.DoesNotContain(page.Comments, comment =>
                comment.Id == targeted.Id ||
                comment.ParentCommentId == targeted.Id);
            using var stream = new MemoryStream();
            document.Save(stream);
            Assert.True(stream.Length > 0);
        }

        [Fact]
        public void ShapeSheetPreservesInterleavedProducerXmlAndRefreshIsTransactional() {
            XNamespace v = "http://schemas.microsoft.com/office/visio/2012/main";
            var source = new XElement(v + "Section", new XAttribute("N", "Actions"),
                new XElement(v + "Row", new XAttribute("N", "One"),
                    new XElement(v + "Cell", new XAttribute("N", "A"), new XAttribute("V", "1")),
                    new XElement(v + "ProducerCellMetadata", new XAttribute("keep", "yes")),
                    new XElement(v + "Cell", new XAttribute("N", "B"), new XAttribute("V", "2"))),
                new XElement(v + "ProducerSectionMetadata", new XAttribute("keep", "yes")),
                new XElement(v + "Row", new XAttribute("N", "Two"),
                    new XElement(v + "Cell", new XAttribute("N", "C"), new XAttribute("V", "3"))));
            var section = new VisioShapeSheetSection(source);
            section.FindRow("One")!.FindCell("A")!.Value = "updated";
            XElement serialized = section.ToXElement();
            Assert.Equal(new[] { "Cell", "ProducerCellMetadata", "Cell" },
                serialized.Elements(v + "Row").First().Elements()
                    .Select(element => element.Name.LocalName));
            Assert.Equal(new[] { "Row", "ProducerSectionMetadata", "Row" },
                serialized.Elements().Select(element => element.Name.LocalName));

            VisioDocument document = VisioDocument.Create();
            VisioPage page = document.AddPage("Graphics", 8.5, 6);
            VisioShape target = page.AddRectangle(3, 3, 1, 1, "Target");
            target.SetShapeData("Status", "Healthy");
            VisioDataGraphic valid = VisioDataGraphic.Create().Badge("Status");
            VisioShape original = Assert.Single(page.AddDataGraphics(target, valid));
            VisioDataGraphic invalid = VisioDataGraphic.Create().Bar("Status",
                minimumValue: 10, maximumValue: 10);
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                page.RefreshDataGraphics(target, invalid));
            Assert.Same(original, Assert.Single(page.GetDataGraphic(target).Shapes));
        }

        private static void InjectVbaProject(string path, byte[] payload) {
            using ZipArchive archive = ZipFile.Open(path, ZipArchiveMode.Update);
            ZipArchiveEntry vba = archive.CreateEntry(
                "visio/macros/project.bin");
            using (Stream output = vba.Open()) output.Write(payload, 0, payload.Length);
            ZipArchiveEntry signature = archive.CreateEntry(
                "visio/macros/signature.bin");
            using (Stream output = signature.Open())
                output.Write(new byte[] { 9, 8, 7, 6 }, 0, 4);
            XNamespace relationships = "http://schemas.openxmlformats.org/package/2006/relationships";
            var vbaRelationships = new XDocument(new XElement(
                relationships + "Relationships",
                new XElement(relationships + "Relationship",
                    new XAttribute("Id", "rIdSignature"),
                    new XAttribute("Type", "http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature"),
                    new XAttribute("Target", "signature.bin")),
                new XElement(relationships + "Relationship",
                    new XAttribute("Id", "rIdExternal"),
                    new XAttribute("Type", "urn:producer:signing-policy"),
                    new XAttribute("Target", "https://example.test/signing-policy"),
                    new XAttribute("TargetMode", "External"))));
            ZipArchiveEntry relationshipEntry = archive.CreateEntry(
                "visio/macros/_rels/project.bin.rels");
            using (Stream output = relationshipEntry.Open())
                vbaRelationships.Save(output);
            UpdateXmlEntry(archive, "visio/_rels/document.xml.rels", document => {
                document.Root!.Add(new XElement(relationships + "Relationship",
                    new XAttribute("Id", "rIdVba"),
                    new XAttribute("Type", "http://schemas.microsoft.com/office/2006/relationships/vbaProject"),
                    new XAttribute("Target", "macros/project.bin")));
            });
            UpdateXmlEntry(archive, "[Content_Types].xml", document => {
                XNamespace types = "http://schemas.openxmlformats.org/package/2006/content-types";
                document.Root!.Add(new XElement(types + "Override",
                    new XAttribute("PartName", "/visio/macros/project.bin"),
                    new XAttribute("ContentType", "application/vnd.ms-office.vbaProject")));
                document.Root!.Add(new XElement(types + "Override",
                    new XAttribute("PartName", "/visio/macros/signature.bin"),
                    new XAttribute("ContentType", "application/vnd.ms-office.vbaProjectSignature")));
            });
        }

        private static void RemovePagesFromStencil(string path) {
            using ZipArchive archive = ZipFile.Open(path, ZipArchiveMode.Update);
            UpdateXmlEntry(archive, "visio/_rels/document.xml.rels", document => {
                XNamespace relationships = "http://schemas.openxmlformats.org/package/2006/relationships";
                document.Root!.Elements(relationships + "Relationship")
                    .Where(element => ((string?)element.Attribute("Type"))?
                        .EndsWith("/pages", StringComparison.Ordinal) == true)
                    .Remove();
            });
            UpdateXmlEntry(archive, "[Content_Types].xml", document => {
                XNamespace types = "http://schemas.openxmlformats.org/package/2006/content-types";
                document.Root!.Elements(types + "Override")
                    .Where(element => ((string?)element.Attribute("PartName"))?
                        .StartsWith("/visio/pages/", StringComparison.OrdinalIgnoreCase) == true)
                    .Remove();
            });
            foreach (ZipArchiveEntry entry in archive.Entries
                .Where(entry => entry.FullName.StartsWith("visio/pages/",
                    StringComparison.OrdinalIgnoreCase)).ToArray()) entry.Delete();
        }

        private static void UpdateXmlEntry(ZipArchive archive, string name,
            Action<XDocument> update) {
            ZipArchiveEntry entry = archive.GetEntry(name)!;
            XDocument document;
            using (Stream input = entry.Open()) document = XDocument.Load(input);
            update(document);
            entry.Delete();
            ZipArchiveEntry replacement = archive.CreateEntry(name);
            using Stream output = replacement.Open();
            document.Save(output);
        }
    }
}
