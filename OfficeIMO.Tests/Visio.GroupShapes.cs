using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioGroupShapes {
        private static string AssetsPath => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets"));
        private static readonly XNamespace VisioNs = XNamespace.Get("http://schemas.microsoft.com/office/visio/2012/main");

        [Fact]
        public void LoadGroupedShapes_PreservesHierarchyAndMasterReferences() {
            string template = Path.Combine(AssetsPath, "VisioTemplates", "DrawingWithJenkinsDiagram.vsdx");

            VisioDocument document = VisioDocument.Load(template);
            VisioShape? groupShape = document.Pages
                .SelectMany(page => page.Shapes)
                .FirstOrDefault(shape => string.Equals(shape.Type, "Group", StringComparison.OrdinalIgnoreCase) || shape.Children.Count > 0);

            Assert.NotNull(groupShape);
            VisioShape group = groupShape!;
            Assert.NotEmpty(group.Children);

            foreach (VisioShape child in group.Children) {
                Assert.Same(group, child.Parent);
            }

            VisioShape? childWithMaster = group.Children.FirstOrDefault(child => !string.IsNullOrEmpty(child.MasterShapeId));
            Assert.NotNull(childWithMaster);
            VisioShape childShape = childWithMaster!;

            Assert.NotNull(childShape.Master);
            Assert.NotNull(childShape.MasterShape);
            Assert.Equal(childShape.MasterShapeId, childShape.MasterShape!.Id);
        }

        [Fact]
        public void GroupedShapesSurviveLoadSaveRoundTrip() {
            string template = Path.Combine(AssetsPath, "VisioTemplates", "DrawingWithJenkinsDiagram.vsdx");
            string output = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Load(template);
            document.Save(output);
            VisioDocument reloaded = VisioDocument.Load(output);

            VisioShape? groupShape = reloaded.Pages
                .SelectMany(page => page.Shapes)
                .FirstOrDefault(shape => string.Equals(shape.Type, "Group", StringComparison.OrdinalIgnoreCase) || shape.Children.Count > 0);

            Assert.NotNull(groupShape);
            Assert.NotEmpty(groupShape!.Children);
        }

        [Fact]
        public void DirectGroupedShapeAuthoringAssignsParentLinksAutomatically() {
            VisioShape group = new("1") {
                Type = "Group",
                PinX = 2,
                PinY = 2,
                Width = 4,
                Height = 3,
                LocPinX = 2,
                LocPinY = 1.5
            };
            VisioShape child = new("2", 1, 1, 1, 1, "Child");
            VisioShape grandChild = new("3", 0.5, 0.5, 0.5, 0.5, "Grandchild");

            child.Children.Add(grandChild);
            group.Children.Add(child);

            Assert.Same(group, child.Parent);
            Assert.Same(child, grandChild.Parent);
        }

        [Fact]
        public void PageShapeCollectionNormalizesGroupedHierarchyForValidation() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Page-1");
            VisioShape group = new("1") {
                Type = "Group",
                PinX = 2,
                PinY = 2,
                Width = 4,
                Height = 3,
                LocPinX = 2,
                LocPinY = 1.5
            };
            VisioShape child = new("2", 1, 1, 1, 1, "Child");
            VisioShape grandChild = new("3", 0.5, 0.5, 0.5, 0.5, "Grandchild");
            child.Children.Add(grandChild);
            group.Children.Add(child);

            page.Shapes.Add(group);

            Assert.Empty(document.Validate());
            Assert.Same(group, child.Parent);
            Assert.Same(child, grandChild.Parent);
        }

        [Fact]
        public void AddingAncestorAsChildThrowsHelpfulError() {
            VisioShape group = new("1") { Type = "Group" };
            VisioShape child = new("2", 1, 1, 1, 1, "Child");
            group.Children.Add(child);

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() => child.Children.Add(group));

            Assert.Contains("cycle", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void AddingShapeToAnotherParentThrowsHelpfulError() {
            VisioShape firstGroup = new("1") { Type = "Group" };
            VisioShape secondGroup = new("2") { Type = "Group" };
            VisioShape child = new("3", 1, 1, 1, 1, "Child");
            firstGroup.Children.Add(child);

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() => secondGroup.Children.Add(child));

            Assert.Contains("another parent", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void AddingChildShapeDirectlyToPageThrowsHelpfulError() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Page-1");
            VisioShape group = new("1") { Type = "Group" };
            VisioShape child = new("2", 1, 1, 1, 1, "Child");
            group.Children.Add(child);

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() => page.Shapes.Add(child));

            Assert.Contains("removed from its parent", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void ReparentShapeMovesTopLevelShapeIntoGroup() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Page-1");
            VisioShape group = new("1") { Type = "Group" };
            VisioShape child = new("2", 1, 1, 1, 1, "Child");
            page.Shapes.Add(group);
            page.Shapes.Add(child);

            page.ReparentShape(child, group);

            Assert.DoesNotContain(child, page.Shapes);
            Assert.Single(group.Children);
            Assert.Same(group, child.Parent);
        }

        [Fact]
        public void ReparentShapeMovesChildBetweenGroupsAndPreservesRequestedIndex() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Page-1");
            VisioShape firstGroup = new("1") { Type = "Group" };
            VisioShape secondGroup = new("2") { Type = "Group" };
            VisioShape existingChild = new("3", 1, 1, 1, 1, "Existing");
            VisioShape movedChild = new("4", 2, 1, 1, 1, "Moved");
            page.Shapes.Add(firstGroup);
            page.Shapes.Add(secondGroup);
            firstGroup.Children.Add(movedChild);
            secondGroup.Children.Add(existingChild);

            page.ReparentShape(movedChild, secondGroup, childIndex: 0);

            Assert.Empty(firstGroup.Children);
            Assert.Equal(new[] { movedChild, existingChild }, secondGroup.Children.ToArray());
            Assert.Same(secondGroup, movedChild.Parent);
        }

        [Fact]
        public void ReparentShapeAppendsWhenMovingWithinSameParent() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Page-1");
            VisioShape group = new("1") { Type = "Group" };
            VisioShape first = new("2", 1, 1, 1, 1, "First");
            VisioShape middle = new("3", 2, 1, 1, 1, "Middle");
            VisioShape last = new("4", 3, 1, 1, 1, "Last");
            page.Shapes.Add(group);
            group.Children.Add(first);
            group.Children.Add(middle);
            group.Children.Add(last);

            page.ReparentShape(first, group, childIndex: -1);

            Assert.Equal(new[] { middle, last, first }, group.Children.ToArray());
            Assert.Same(group, first.Parent);
        }

        [Fact]
        public void ReplacingChildWithDuplicateLeavesExistingParentLinkIntact() {
            VisioShape group = new("1") { Type = "Group" };
            VisioShape first = new("2", 1, 1, 1, 1, "First");
            VisioShape second = new("3", 2, 1, 1, 1, "Second");
            group.Children.Add(first);
            group.Children.Add(second);

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() => group.Children[0] = second);

            Assert.Contains("already a child", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Same(group, first.Parent);
            Assert.Same(group, second.Parent);
            Assert.Equal(new[] { first, second }, group.Children.ToArray());
        }

        [Fact]
        public void UngroupShapePromotesChildrenIntoFormerParentSlot() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Page-1");
            VisioShape before = new("1", 1, 1, 1, 1, "Before");
            VisioShape group = new("2") { Type = "Group" };
            VisioShape childOne = new("3", 2, 1, 1, 1, "Child-1");
            VisioShape childTwo = new("4", 3, 1, 1, 1, "Child-2");
            VisioShape after = new("5", 4, 1, 1, 1, "After");
            group.Children.Add(childOne);
            group.Children.Add(childTwo);
            page.Shapes.Add(before);
            page.Shapes.Add(group);
            page.Shapes.Add(after);

            IReadOnlyList<VisioShape> promoted = page.UngroupShape(group);

            Assert.Equal(new[] { childOne, childTwo }, promoted.ToArray());
            Assert.Equal(new[] { before, childOne, childTwo, after }, page.Shapes.ToArray());
            Assert.Empty(group.Children);
            Assert.Null(childOne.Parent);
            Assert.Null(childTwo.Parent);
        }

        [Fact]
        public void UngroupShapeWorksForNestedGroups() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Page-1");
            VisioShape outerGroup = new("1") { Type = "Group" };
            VisioShape innerGroup = new("2") { Type = "Group" };
            VisioShape child = new("3", 1, 1, 1, 1, "Child");
            innerGroup.Children.Add(child);
            outerGroup.Children.Add(innerGroup);
            page.Shapes.Add(outerGroup);

            page.UngroupShape(innerGroup);

            Assert.Equal(new[] { child }, outerGroup.Children.ToArray());
            Assert.Same(outerGroup, child.Parent);
            Assert.Empty(innerGroup.Children);
        }

        [Fact]
        public void ReconnectConnectorStartCanRetargetPromotedChildAfterUngroup() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Page-1");
            VisioShape group = new("1") { Type = "Group", Width = 2, Height = 2, LocPinX = 1, LocPinY = 1 };
            VisioShape promotedChild = new("2", 2, 2, 2, 2, "Promoted");
            VisioShape target = new("3", 6, 2, 2, 2, "Target");
            group.Children.Add(promotedChild);
            page.Shapes.Add(group);
            page.Shapes.Add(target);
            VisioConnector connector = page.AddConnector(group, target, ConnectorKind.Straight);

            page.UngroupShape(group);
            page.ReconnectConnectorStart(connector, promotedChild, VisioSide.Right);

            Assert.Same(promotedChild, connector.From);
            Assert.Equal(promotedChild.Width, connector.FromConnectionPoint!.X, 5);
            Assert.Empty(document.Validate());

            document.Save();

            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XDocument pageXml = ReadPageXml(filePath);
            XElement beginConnect = pageXml.Root!
                .Element(ns + "Connects")!
                .Elements(ns + "Connect")
                .First(connect => (string?)connect.Attribute("FromSheet") == connector.Id && (string?)connect.Attribute("FromCell") == "BeginX");
            Assert.Equal(promotedChild.Id, (string?)beginConnect.Attribute("ToSheet"));
        }

        [Fact]
        public void ReconnectConnectorCanUpdateBothEndpointsAndSideGlue() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Page-1");
            VisioShape left = new("1", 1, 1, 2, 2, "Left");
            VisioShape middle = new("2", 4, 1, 2, 2, "Middle");
            VisioShape right = new("3", 7, 1, 2, 2, "Right");
            page.Shapes.Add(left);
            page.Shapes.Add(middle);
            page.Shapes.Add(right);
            VisioConnector connector = page.AddConnector(left, middle, ConnectorKind.Straight);

            page.ReconnectConnector(connector, middle, right, VisioSide.Right, VisioSide.Left);

            Assert.Same(middle, connector.From);
            Assert.Same(right, connector.To);
            Assert.Equal(middle.Width, connector.FromConnectionPoint!.X, 5);
            Assert.Equal(0, connector.ToConnectionPoint!.X, 5);
        }

        [Fact]
        public void ReconnectConnectorRejectsShapesOutsideThePage() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Page-1");
            VisioShape start = new("1", 1, 1, 2, 2, "Start");
            VisioShape end = new("2", 4, 1, 2, 2, "End");
            VisioShape external = new("3", 7, 1, 2, 2, "External");
            page.Shapes.Add(start);
            page.Shapes.Add(end);
            VisioConnector connector = page.AddConnector(start, end, ConnectorKind.Straight);

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                page.ReconnectConnectorEnd(connector, external, VisioSide.Left));

            Assert.Contains("not part of this page", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void RetargetConnectorsCanMigrateGroupEdgesToPromotedChild() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Page-1");
            VisioShape source = new("1", 1, 1, 2, 2, "Source");
            VisioShape group = new("2") { Type = "Group", Width = 2, Height = 2, LocPinX = 1, LocPinY = 1 };
            VisioShape promotedChild = new("3", 4, 1, 2, 2, "Promoted");
            VisioShape target = new("4", 7, 1, 2, 2, "Target");
            group.Children.Add(promotedChild);
            page.Shapes.Add(source);
            page.Shapes.Add(group);
            page.Shapes.Add(target);
            VisioConnector incoming = page.AddConnector(source, group, ConnectorKind.Straight);
            VisioConnector outgoing = page.AddConnector(group, target, ConnectorKind.Straight);
            VisioConnector unaffected = page.AddConnector(source, target, ConnectorKind.Straight);

            page.UngroupShape(group);
            IReadOnlyList<VisioConnector> updated = page.RetargetConnectors(group, promotedChild, fromSide: VisioSide.Right, toSide: VisioSide.Left);

            Assert.Equal(new[] { incoming, outgoing }, updated.ToArray());
            Assert.Same(promotedChild, incoming.To);
            Assert.Same(promotedChild, outgoing.From);
            Assert.Equal(0, incoming.ToConnectionPoint!.X, 5);
            Assert.Equal(promotedChild.Width, outgoing.FromConnectionPoint!.X, 5);
            Assert.Same(source, unaffected.From);
            Assert.Same(target, unaffected.To);
            Assert.Empty(document.Validate());
        }

        [Fact]
        public void RetargetConnectorsCanLimitUpdatesToStartEndpoints() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Page-1");
            VisioShape oldShape = new("1", 1, 1, 2, 2, "Old");
            VisioShape replacement = new("2", 4, 1, 2, 2, "Replacement");
            VisioShape target = new("3", 7, 1, 2, 2, "Target");
            page.Shapes.Add(oldShape);
            page.Shapes.Add(replacement);
            page.Shapes.Add(target);
            VisioConnector startMatch = page.AddConnector(oldShape, target, ConnectorKind.Straight);
            VisioConnector endMatch = page.AddConnector(target, oldShape, ConnectorKind.Straight);

            IReadOnlyList<VisioConnector> updated = page.RetargetConnectors(oldShape, replacement, VisioConnectorEndpointScope.Start, fromSide: VisioSide.Bottom);

            Assert.Equal(new[] { startMatch }, updated.ToArray());
            Assert.Same(replacement, startMatch.From);
            Assert.Equal(0, startMatch.FromConnectionPoint!.Y, 5);
            Assert.Same(oldShape, endMatch.To);
            Assert.Null(endMatch.ToConnectionPoint);
        }

        [Fact]
        public void RetargetConnectorsReturnsEmptyWhenNoConnectorsReferenceShape() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Page-1");
            VisioShape original = new("1", 1, 1, 2, 2, "Original");
            VisioShape replacement = new("2", 4, 1, 2, 2, "Replacement");
            VisioShape source = new("3", 7, 1, 2, 2, "Source");
            VisioShape target = new("4", 10, 1, 2, 2, "Target");
            page.Shapes.Add(original);
            page.Shapes.Add(replacement);
            page.Shapes.Add(source);
            page.Shapes.Add(target);
            VisioConnector connector = page.AddConnector(source, target, ConnectorKind.Straight);

            IReadOnlyList<VisioConnector> updated = page.RetargetConnectors(original, replacement);

            Assert.Empty(updated);
            Assert.Same(source, connector.From);
            Assert.Same(target, connector.To);
        }

        [Fact]
        public void ParseGroupWithoutChildren_AllowsEmptyGroups() {
            XElement groupElement = new XElement(VisioNs + "Shape",
                new XAttribute("ID", "1"),
                new XAttribute("Type", "Group"));

            VisioShape shape = InvokeParseShape(groupElement);

            Assert.Equal("1", shape.Id);
            Assert.Empty(shape.Children);
        }

        [Fact]
        public void ApplyMasterReferences_IgnoresMissingMasterDefinitions() {
            XElement shapeElement = new XElement(VisioNs + "Shape",
                new XAttribute("ID", "1"),
                new XAttribute("Master", "999"),
                new XAttribute("MasterShape", "2"));

            VisioShape shape = InvokeParseShape(shapeElement);
            Dictionary<string, VisioMaster> masters = new();

            InvokeApplyMasterReferences(shape, shapeElement, masters);

            Assert.Null(shape.Master);
            Assert.Null(shape.MasterShape);
        }

        [Fact]
        public void ApplyMasterReferences_UsesMatchingMasterChildWhenGroupedChildrenLackMasterShapeIds() {
            XElement shapeElement = new XElement(VisioNs + "Shape",
                new XAttribute("ID", "1"),
                new XAttribute("Master", "100"),
                new XAttribute("Type", "Group"),
                new XElement(VisioNs + "Shapes",
                    new XElement(VisioNs + "Shape", new XAttribute("ID", "2")),
                    new XElement(VisioNs + "Shape", new XAttribute("ID", "3"))));

            VisioShape shape = InvokeParseShape(shapeElement);

            VisioShape masterRoot = new("10") {
                Type = "Group",
                Width = 9,
                Height = 9,
                LocPinX = 4.5,
                LocPinY = 4.5
            };
            masterRoot.Children.Add(new VisioShape("11") { Width = 2, Height = 1, LocPinX = 1, LocPinY = 0.5 });
            masterRoot.Children.Add(new VisioShape("12") { Width = 4, Height = 3, LocPinX = 2, LocPinY = 1.5 });

            Dictionary<string, VisioMaster> masters = new() {
                ["100"] = new VisioMaster("100", "GroupedMaster", masterRoot)
            };

            InvokeApplyMasterReferences(shape, shapeElement, masters);

            Assert.Equal(2, shape.Children[0].Width);
            Assert.Equal(1, shape.Children[0].Height);
            Assert.Equal(1, shape.Children[0].LocPinX);
            Assert.Equal(0.5, shape.Children[0].LocPinY);

            Assert.Equal(4, shape.Children[1].Width);
            Assert.Equal(3, shape.Children[1].Height);
            Assert.Equal(2, shape.Children[1].LocPinX);
            Assert.Equal(1.5, shape.Children[1].LocPinY);
        }

        [Fact]
        public void ParseShape_AllowsNestingUpToLimit() {
            int maxDepth = GetMaxShapeDepth();
            XElement root = CreateNestedGroup(maxDepth);

            VisioShape shape = InvokeParseShape(root);

            Assert.Equal("0", shape.FindDescendantById("0")?.Id);
        }

        [Fact]
        public void ParseShape_ThrowsWhenNestingExceedsLimit() {
            int maxDepth = GetMaxShapeDepth();
            XElement root = CreateNestedGroup(maxDepth + 1);

            Assert.Throws<InvalidOperationException>(() => InvokeParseShape(root));
        }

        private static VisioShape InvokeParseShape(XElement shapeElement) {
            MethodInfo parseMethod = typeof(VisioDocument).GetMethod("ParseShape", BindingFlags.NonPublic | BindingFlags.Static)!;
            try {
                return (VisioShape)parseMethod.Invoke(null, new object?[] { shapeElement, VisioNs, null, 0 })!;
            } catch (TargetInvocationException ex) when (ex.InnerException != null) {
                throw ex.InnerException;
            }
        }

        private static void InvokeApplyMasterReferences(VisioShape shape, XElement shapeElement, Dictionary<string, VisioMaster> masters) {
            MethodInfo applyMethod = typeof(VisioDocument).GetMethod("ApplyMasterReferences", BindingFlags.NonPublic | BindingFlags.Static)!;
            applyMethod.Invoke(null, new object?[] { shape, shapeElement, VisioNs, masters, null, null });
        }

        private static int GetMaxShapeDepth() {
            FieldInfo field = typeof(VisioDocument).GetField("MaxShapeNestingDepth", BindingFlags.NonPublic | BindingFlags.Static)!;
            return (int)field.GetValue(null)!;
        }

        private static XElement CreateNestedGroup(int depth) {
            XElement element = new XElement(VisioNs + "Shape",
                new XAttribute("ID", depth.ToString()),
                new XAttribute("Type", "Group"));

            if (depth > 0) {
                XElement child = CreateNestedGroup(depth - 1);
                XElement shapesContainer = new XElement(VisioNs + "Shapes", child);
                element.Add(shapesContainer);
            }

            return element;
        }

        private static XDocument ReadPageXml(string vsdxPath) {
            using FileStream stream = File.Open(vsdxPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using ZipArchive archive = new(stream, ZipArchiveMode.Read);
            using Stream pageStream = archive.GetEntry("visio/pages/page1.xml")!.Open();
            return XDocument.Load(pageStream);
        }
    }
}
