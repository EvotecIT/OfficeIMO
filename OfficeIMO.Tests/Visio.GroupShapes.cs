using System;
using System.Collections.Generic;
using System.IO;
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
    }
}
