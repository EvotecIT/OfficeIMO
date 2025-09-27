using System;
using System.IO;
using System.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioGroupShapes {
        private static string AssetsPath => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets"));

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
    }
}
