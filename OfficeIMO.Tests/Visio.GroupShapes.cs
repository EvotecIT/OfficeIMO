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
