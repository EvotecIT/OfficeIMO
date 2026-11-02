using System;
using System.IO;
using System.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Stencils;
using Xunit;
using Color = OfficeIMO.OfficeColor;

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
                PinX = OfficeLength.Inches(2),
                PinY = OfficeLength.Inches(2),
                Width = OfficeLength.Inches(4),
                Height = OfficeLength.Inches(3),
                LocPinX = OfficeLength.Inches(2),
                LocPinY = OfficeLength.Inches(1.5
            )};
            VisioShape child = new("child", OfficeLength.Inches(1), OfficeLength.Inches(1), OfficeLength.Inches(1), OfficeLength.Inches(1), "Child");
            VisioShape grandChild = new("grandchild", OfficeLength.Inches(0.5), OfficeLength.Inches(0.5), OfficeLength.Inches(0.4), OfficeLength.Inches(0.4), "Grandchild");

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
            VisioPage page = document.AddPage("Tagged", OfficeLength.Inches(11), OfficeLength.Inches(8.5));
            VisioShape intake = page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "intake", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(6, page.DefaultUnit)), "Intake");
            VisioShape review = page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "review", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(6, page.DefaultUnit)), "Review");
            VisioShape archive = page.AddStencilShape(VisioStencils.Flowchart.Get("data"), "archive", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(6, page.DefaultUnit)), "Archive");
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
        public void ShapeDataValueQueriesMatchLegacyDataDictionaryValues() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Legacy Data");
            VisioShape web = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "Web");
            VisioShape api = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "API");
            web.Data["Owner"] = "Ops";
            api.SetShapeData("Owner", "Platform");

            Assert.Same(web, Assert.Single(page.ShapesWithData("Owner", "Ops")));
            Assert.Same(web, Assert.Single(page.ShapesWithData("Owner", value => string.Equals(value, "Ops", StringComparison.OrdinalIgnoreCase))));
            Assert.Same(web, page.SelectWithData("Owner", "ops").Single());
        }

        [Fact]
        public void ConnectorQueriesFindNeighborsAndSupportBulkStyling() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Connectors");
            VisioShape source = page.AddStencilShape(VisioStencils.BasicShapes.Get("rectangle"), "source", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5, page.DefaultUnit)), "Source");
            VisioShape middle = page.AddStencilShape(VisioStencils.BasicShapes.Get("rectangle"), "middle", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5, page.DefaultUnit)), "Middle");
            VisioShape target = page.AddStencilShape(VisioStencils.BasicShapes.Get("rectangle"), "target", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5, page.DefaultUnit)), "Target");
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
        public void GeometryQueriesSelectIntersectingAndContainedShapes() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Geometry");
            VisioShape zone = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(6, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), "Zone");
            VisioShape inside = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "Inside");
            VisioShape edge = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(7.4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "Edge");
            VisioShape outside = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(9.5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "Outside");
            VisioShape overlap = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2.5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1.4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "Overlap");

            Assert.Equal(new[] { "Edge", "Inside" }, page.ShapesContainedIn(zone).Select(shape => shape.Text).OrderBy(text => text).ToArray());
            Assert.Equal(new[] { "Edge", "Inside", "Overlap" }, page.ShapesIntersecting(zone).Select(shape => shape.Text).OrderBy(text => text).ToArray());
            Assert.Equal(new[] { "Edge", "Inside", "Zone" }, page.ShapesContainedIn(new OfficeIMO.Visio.VisioShapeBounds(OfficeLength.Inches(1.9), OfficeLength.Inches(1.9), OfficeLength.Inches(8.1), OfficeLength.Inches(6.1))).Select(shape => shape.Text).OrderBy(text => text).ToArray());

            page.SelectContainedIn(zone)
                .ShapeData("Zone", "Trusted")
                .Fill(Color.LightGreen);

            Assert.Equal("Trusted", inside.GetShapeDataValue("Zone"));
            Assert.Equal("Trusted", edge.GetShapeDataValue("Zone"));
            Assert.Null(overlap.GetShapeDataValue("Zone"));
            Assert.Equal(Color.LightGreen, inside.FillColor);
            Assert.Equal(Color.White, outside.FillColor);
        }

        [Fact]
        public void GeometryQueriesResolveGroupedChildBoundsInPageCoordinates() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Grouped Geometry");
            VisioShape group = new("group") {
                Type = "Group",
                PinX = OfficeLength.Inches(5),
                PinY = OfficeLength.Inches(4),
                Width = OfficeLength.Inches(4),
                Height = OfficeLength.Inches(3),
                LocPinX = OfficeLength.Inches(2),
                LocPinY = OfficeLength.Inches(1.5
            )};
            VisioShape child = new("child", OfficeLength.Inches(1), OfficeLength.Inches(1), OfficeLength.Inches(1), OfficeLength.Inches(1), "Child");
            group.Children.Add(child);
            page.Shapes.Add(group);

            OfficeIMO.Visio.VisioShapeBounds pageBounds = new(OfficeLength.Inches(3.25), OfficeLength.Inches(2.75), OfficeLength.Inches(4.75), OfficeLength.Inches(4.25));
            OfficeIMO.Visio.VisioShapeBounds localBounds = new(OfficeLength.Inches(0.5), OfficeLength.Inches(0.5), OfficeLength.Inches(1.5), OfficeLength.Inches(1.5));

            Assert.Contains(child, page.ShapesIntersecting(pageBounds));
            Assert.Contains(child, page.ShapesContainedIn(pageBounds));
            Assert.DoesNotContain(child, page.ShapesIntersecting(localBounds));
        }

        [Fact]
        public void ShapeDataPredicateQueriesFindMatchingValues() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Predicates");
            VisioShape web = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "Web");
            VisioShape api = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "API");
            VisioShape database = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(6, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "Database");
            web.SetShapeData("Risk", "2");
            api.SetShapeData("Risk", "4");
            database.SetShapeData("Risk", "5");

            VisioShapeSelection highRisk = page.SelectWithShapeData("Risk", value => int.TryParse(value, out int risk) && risk >= 4)
                .Stroke(Color.Red, 0.025);

            Assert.Equal(new[] { "API", "Database" }, highRisk.Select(shape => shape.Text).OrderBy(text => text).ToArray());
            Assert.Equal(Color.Black, web.LineColor);
            Assert.Equal(Color.Red, api.LineColor);
            Assert.Equal(Color.Red, database.LineColor);
        }

        [Fact]
        public void ConnectedComponentAndPathQueriesSupportGraphEditing() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Path");
            VisioShape source = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "Source");
            VisioShape gateway = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(3, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "Gateway");
            VisioShape api = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "API");
            VisioShape database = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(7, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "Database");
            VisioShape isolated = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(7, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "Isolated");
            page.AddConnector(source, gateway);
            page.AddConnector(gateway, api);
            page.AddConnector(api, database);

            Assert.Equal(new[] { "API", "Database", "Gateway", "Source" }, page.ConnectedComponent(source).Select(shape => shape.Text).OrderBy(text => text).ToArray());
            Assert.Equal(Array.Empty<string>(), page.ConnectedComponent(isolated, includeStart: false).Select(shape => shape.Id).ToArray());
            Assert.Equal(new[] { "Source", "Gateway", "API", "Database" }, page.PathBetween(source, database).Select(shape => shape.Text).ToArray());
            Assert.Equal(new[] { "Gateway", "API" }, page.PathBetween(source, database, includeEndpoints: false).Select(shape => shape.Text).ToArray());
            Assert.Empty(page.PathBetween(source, isolated));

            page.SelectConnectedComponent(source)
                .ShapeData("Component", "Payments");

            Assert.Equal("Payments", database.GetShapeDataValue("Component"));
            Assert.Null(isolated.GetShapeDataValue("Component"));
        }

        [Fact]
        public void ConnectorQueriesRejectShapesFromOtherPagesEvenWhenIdsMatch() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage first = document.AddPage("First");
            VisioShape firstSource = first.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, first.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, first.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, first.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, first.DefaultUnit)), "Source");
            VisioShape firstTarget = first.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, first.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, first.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, first.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, first.DefaultUnit)), "Target");
            first.AddConnector(firstSource, firstTarget);
            VisioPage second = document.AddPage("Second");
            VisioShape duplicateId = second.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, second.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, second.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, second.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, second.DefaultUnit)), "Duplicate");

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
            page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "one", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5, page.DefaultUnit)), "One").Data["Layer"] = "A";
            page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "two", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5, page.DefaultUnit)), "Two").Data["Layer"] = "A";

            VisioShapeSelection selection = page.SelectShapes(shape => shape.Data.ContainsKey("Layer"));
            page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "three", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5, page.DefaultUnit)), "Three").Data["Layer"] = "A";

            selection.Data("Reviewed", "Yes").MoveBy(1, -1).Size(OfficeLength.Inches(1.5), OfficeLength.Inches(0.75));

            Assert.Equal(2, selection.Count);
            Assert.Equal(3, page.ShapesWithData("Layer", "A").Count);
            Assert.Equal(2, page.ShapesWithData("Reviewed", "Yes").Count);
            Assert.All(selection, shape => Assert.Equal(1.5, shape.Width.ToInches()));
            Assert.All(selection, shape => Assert.Equal(0.75, shape.Height.ToInches()));
        }
    }
}
