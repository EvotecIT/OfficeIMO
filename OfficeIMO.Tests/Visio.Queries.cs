using System;
using System.IO;
using System.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Stencils;
using Xunit;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Tests {
    public class VisioQueriesTests {
        [Fact]
        public void AllShapesAndFindShapeByIdIncludeGroupedChildren() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Query");
            VisioShape group = new("group") {
                Name = "Container",
                NameU = "Container",
                Type = "Group",
                PinX = 2,
                PinY = 2,
                Width = 4,
                Height = 3,
                LocPinX = 2,
                LocPinY = 1.5
            };
            VisioShape child = new("child", 1, 1, 1, 1, "Child") {
                Name = "Nested child",
                NameU = "NestedChild"
            };
            VisioShape grandChild = new("grandchild", 0.5, 0.5, 0.4, 0.4, "Grandchild");

            child.Children.Add(grandChild);
            group.Children.Add(child);
            page.Shapes.Add(group);

            Assert.Equal(new[] { "group", "child", "grandchild" }, page.AllShapes().Select(shape => shape.Id).ToArray());
            Assert.Same(grandChild, page.FindShapeById("grandchild"));
            Assert.Same(child, page.ShapesByName("Nested child").Single());
            Assert.Same(child, page.ShapesByNameU("nestedchild").Single());
        }

        [Fact]
        public void ShapeSelectionCanFindTaggedStencilShapesAndBulkEditThem() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Tagged", 11, 8.5);
            VisioShape intake = page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "intake", 2, 6, "Intake");
            VisioShape review = page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "review", 5, 6, "Review");
            VisioShape archive = page.AddStencilShape(VisioStencils.Flowchart.Get("data"), "archive", 8, 6, "Archive");
            intake.Data["Owner"] = "Ops";
            review.Data["Owner"] = "Ops";
            archive.Data["Owner"] = "Records";

            VisioShapeSelection opsShapes = page.SelectWithData("Owner", "ops")
                .Fill(Color.LightBlue)
                .Stroke(Color.DodgerBlue, 0.02)
                .Text(shape => shape.Text + " owned");

            Assert.Equal(2, opsShapes.Count);
            Assert.All(opsShapes, shape => Assert.Equal(Color.LightBlue, shape.FillColor));
            Assert.All(opsShapes, shape => Assert.Equal(Color.DodgerBlue, shape.LineColor));
            Assert.All(opsShapes, shape => Assert.EndsWith("owned", shape.Text));
            Assert.Same(archive, page.ShapesByMaster("Data").Single());

            document.Save();

            Assert.Empty(VisioValidator.Validate(filePath));
            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Equal(2, loaded.Pages[0].ShapesWithData("Owner", "Ops").Count);
            Assert.Equal(2, loaded.Pages[0].ShapesContainingText("owned").Count);
        }

        [Fact]
        public void ConnectorQueriesFindNeighborsAndSupportBulkStyling() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Connectors");
            VisioShape source = page.AddStencilShape(VisioStencils.BasicShapes.Get("rectangle"), "source", 2, 5, "Source");
            VisioShape middle = page.AddStencilShape(VisioStencils.BasicShapes.Get("rectangle"), "middle", 5, 5, "Middle");
            VisioShape target = page.AddStencilShape(VisioStencils.BasicShapes.Get("rectangle"), "target", 8, 5, "Target");
            VisioConnector first = page.AddConnector(source, middle, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left);
            VisioConnector second = page.AddConnector(middle, target, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left);

            page.SelectOutgoingConnectors(middle)
                .Kind(ConnectorKind.RightAngle)
                .Stroke(Color.Orange, 0.03, pattern: 2)
                .EndArrow(EndArrow.Triangle)
                .Label("next");

            Assert.Same(first, page.IncomingConnectors(middle).Single());
            Assert.Same(second, page.OutgoingConnectors(middle).Single());
            Assert.Equal(new[] { "source", "target" }, page.ConnectedShapes(middle).Select(shape => shape.Id).OrderBy(id => id).ToArray());
            Assert.Equal(2, page.ConnectedConnectors(middle).Count);
            Assert.Equal(ConnectorKind.RightAngle, second.Kind);
            Assert.Equal(Color.Orange, second.LineColor);
            Assert.Equal(0.03, second.LineWeight);
            Assert.Equal(2, second.LinePattern);
            Assert.Equal(EndArrow.Triangle, second.EndArrow);
            Assert.Equal("next", second.Label);
            Assert.Equal(Color.Black, first.LineColor);
        }

        [Fact]
        public void ConnectorQueriesRejectShapesFromOtherPagesEvenWhenIdsMatch() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage first = document.AddPage("First");
            VisioShape firstSource = first.AddRectangle(2, 2, 1, 1, "Source");
            VisioShape firstTarget = first.AddRectangle(4, 2, 1, 1, "Target");
            first.AddConnector(firstSource, firstTarget);
            VisioPage second = document.AddPage("Second");
            VisioShape duplicateId = second.AddRectangle(2, 2, 1, 1, "Duplicate");

            Assert.Equal(firstSource.Id, duplicateId.Id);
            Assert.Throws<InvalidOperationException>(() => first.OutgoingConnectors(duplicateId));
            Assert.Throws<InvalidOperationException>(() => first.IncomingConnectors(duplicateId));
            Assert.Throws<InvalidOperationException>(() => first.ConnectedConnectors(duplicateId));
            Assert.Throws<InvalidOperationException>(() => first.ConnectedShapes(duplicateId));
        }

        [Fact]
        public void GenericSelectionsMaterializeStableEditableSnapshots() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Snapshot");
            page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "one", 2, 5, "One").Data["Layer"] = "A";
            page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "two", 5, 5, "Two").Data["Layer"] = "A";

            VisioShapeSelection selection = page.SelectShapes(shape => shape.Data.ContainsKey("Layer"));
            page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "three", 8, 5, "Three").Data["Layer"] = "A";

            selection.Data("Reviewed", "Yes").MoveBy(1, -1).Size(1.5, 0.75);

            Assert.Equal(2, selection.Count);
            Assert.Equal(3, page.ShapesWithData("Layer", "A").Count);
            Assert.Equal(2, page.ShapesWithData("Reviewed", "Yes").Count);
            Assert.All(selection, shape => Assert.Equal(1.5, shape.Width));
            Assert.All(selection, shape => Assert.Equal(0.75, shape.Height));
        }
    }
}
