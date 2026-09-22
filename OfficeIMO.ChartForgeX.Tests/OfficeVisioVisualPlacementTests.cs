using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using ChartForgeX.Topology;
using ChartForgeX.Primitives;
using ChartForgeX.VisualArtifacts;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.ChartForgeX.Tests;

public sealed partial class OfficeVisioVisualIntegrationTests {
    [Fact]
    public void NarrowPreservedPageReportsOmittedNativeTitle() {
        var envelope = PlacementEnvelope();
        envelope.Title = "Service";
        envelope.Width = 80;
        envelope.Nodes.RemoveAt(1);
        envelope.Edges.Clear();
        var node = envelope.Nodes[0];
        node.Ports.Clear(); node.X = 10; node.Y = 200; node.Width = 60; node.Height = 80;
        var result = envelope.ToOfficeVisio(new OfficeVisioVisualOptions { PixelsPerInch = 100 });
        Assert.Contains(result.Report.Diagnostics, item => item.Code == OfficeVisioVisualDiagnosticCode.TitleNotProjected);
        Assert.Throws<OfficeVisioVisualFidelityException>(() => envelope.ToOfficeVisio(new OfficeVisioVisualOptions { PixelsPerInch = 100, RequireLossless = true }));
    }

    [Theory]
    [InlineData(OfficeVisioVisualLayoutMode.Auto)]
    [InlineData(OfficeVisioVisualLayoutMode.Preserve)]
    public void EmptyUnpositionedGroupDoesNotDiscardNodePlacement(OfficeVisioVisualLayoutMode mode) {
        var envelope = PlacementEnvelope();
        envelope.Title = "Services";
        envelope.Groups.Add(new VisualArtifactInterchangeGroup { Id = "empty", Label = "Empty", Kind = "TopologyGroup", Role = VisualArtifactInterchangeGroupRole.TopologyGroup, Topology = new VisualArtifactInterchangeTopologyGroup() });
        var result = envelope.ToOfficeVisio(new OfficeVisioVisualOptions { PixelsPerInch = 100, LayoutMode = mode });
        Assert.DoesNotContain(result.Report.Diagnostics, item => item.Code == OfficeVisioVisualDiagnosticCode.LayoutRecomputed);
        Assert.Contains(result.Report.Diagnostics, item => item.Code == OfficeVisioVisualDiagnosticCode.GroupNotProjected);
        Assert.Equal(2, result.Page.Shapes.Single(item => item.Id == "api").PinX, 6);
    }

    [Fact]
    public void PreparedBoundsAndNamedAttachmentsSurviveSaveAndLoad() {
        var envelope = PlacementEnvelope();
        var result = envelope.ToOfficeVisio(new OfficeVisioVisualOptions { PixelsPerInch = 100 });
        Assert.DoesNotContain(result.Report.Diagnostics, item => item.Code == OfficeVisioVisualDiagnosticCode.LayoutRecomputed);
        void Check(VisioPage page) {
            Assert.Equal(10, page.Width, 6);
            Assert.Equal(8, page.Height, 6);
            var node = page.Shapes.Single(item => item.Id == "api");
            Assert.Equal(2, node.PinX, 6);
            Assert.Equal(6.5, node.PinY, 6);
            Assert.Equal(2, node.Width, 6);
            Assert.Equal(1, node.Height, 6);
            var connector = Assert.Single(page.Connectors);
            Assert.Equal(VisioConnectorRerouteBehavior.Never, connector.RerouteBehavior);
            Assert.Equal(2, connector.FromConnectionPoint!.X, 6);
            Assert.Equal(.75, connector.FromConnectionPoint.Y, 6);
            Assert.Contains(connector.Waypoints, point => Math.Abs(point.X - 4) < .000001 && Math.Abs(point.Y - 6.75) < .000001);
        }
        Check(result.Page);
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
        try {
            result.Document.Save(path);
            Assert.Empty(VisioValidator.Validate(path));
            Check(VisioDocument.Load(path).Pages[0]);
        } finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void PreservedDiagonalRouteStaysStraightWithLosslessPolicy() {
        var envelope = PlacementEnvelope();
        envelope.Edges[0].Topology!.Waypoints.Clear();
        var result = envelope.ToOfficeVisio(new OfficeVisioVisualOptions { PixelsPerInch = 100, RequireLossless = true });
        void Check(VisioPage page) {
            var connector = Assert.Single(page.Connectors);
            Assert.Equal(ConnectorKind.Straight, connector.Kind);
            Assert.Empty(connector.Waypoints);
            XNamespace ns = "http://www.w3.org/2000/svg";
            var svg = XDocument.Parse(page.ToSvg());
            var group = svg.Descendants(ns + "g").Single(item => (string?)item.Attribute("data-visio-connector-id") == connector.Id);
            string pathData = (string)group.Elements(ns + "path").Single(item => item.Attribute("data-officeimo-connector-arrow") == null).Attribute("d")!;
            Assert.Equal(1, pathData.Count(value => value == 'M'));
            Assert.Equal(1, pathData.Count(value => value == 'L'));
        }
        Check(result.Page);
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
        try {
            result.Document.Save(path);
            Assert.Empty(VisioValidator.Validate(path));
            Check(VisioDocument.Load(path).Pages[0]);
        } finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void PreservedTitleUsesClearHeaderSpaceOrReportsOmission() {
        var envelope = PlacementEnvelope();
        envelope.Title = "Services";
        var result = envelope.ToOfficeVisio(new OfficeVisioVisualOptions { PixelsPerInch = 100 });
        var title = result.Page.Shapes.Single(shape => shape.Text == "Services");
        var node = result.Page.Shapes.Single(shape => shape.Id == "api");
        Assert.True(title.PinY - title.Height / 2 > node.PinY + node.Height / 2);
        envelope.Nodes[0].Y = 0;
        var crowded = envelope.ToOfficeVisio();
        Assert.DoesNotContain(crowded.Page.Shapes, shape => shape.Text == "Services");
        Assert.Contains(crowded.Report.Diagnostics, item => item.Code == OfficeVisioVisualDiagnosticCode.TitleNotProjected);
        Assert.Throws<OfficeVisioVisualFidelityException>(() => envelope.ToOfficeVisio(new OfficeVisioVisualOptions { RequireLossless = true }));
    }

    [Fact]
    public void ExplicitReflowAndFidelityPolicyAreHonored() {
        var envelope = PlacementEnvelope();
        var reflow = envelope.ToOfficeVisio(new OfficeVisioVisualOptions { LayoutMode = OfficeVisioVisualLayoutMode.Reflow });
        Assert.Contains(reflow.Report.Diagnostics, item => item.Code == OfficeVisioVisualDiagnosticCode.LayoutRecomputed);
        var options = new OfficeVisioVisualOptions { LayoutMode = OfficeVisioVisualLayoutMode.Reflow };
        options.RejectedDiagnostics.Add(OfficeVisioVisualDiagnosticCode.LayoutRecomputed);
        var error = Assert.Throws<OfficeVisioVisualFidelityException>(() => envelope.ToOfficeVisio(options));
        Assert.Contains(error.Report.Diagnostics, item => item.Code == OfficeVisioVisualDiagnosticCode.LayoutRecomputed);
        envelope.Nodes[0].Width = null;
        Assert.Throws<NotSupportedException>(() => envelope.ToOfficeVisio(new OfficeVisioVisualOptions { LayoutMode = OfficeVisioVisualLayoutMode.Preserve }));
    }

    [Fact]
    public void PreserveRejectsSequenceInsteadOfSilentlyReflowing() {
        Assert.Throws<NotSupportedException>(() => SequenceEnvelope("sequence").ToOfficeVisio(
            new OfficeVisioVisualOptions { LayoutMode = OfficeVisioVisualLayoutMode.Preserve }));
    }

    [Fact]
    public void BookPreservesInputOrderAndDistinctPagesThroughSave() {
        var first = PlacementEnvelope(); var second = PlacementEnvelope();
        first.Title = second.Title = "Services";
        var book = new[] { first, second }.ToOfficeVisioBook();
        Assert.Equal(new[] { "Services", "Services (2)" }, book.Document.Pages.Select(page => page.Name));
        Assert.All(book.Pages, page => Assert.Same(book.Document, page.Document));
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
        try {
            book.Document.Save(path);
            Assert.Empty(VisioValidator.Validate(path));
            Assert.Equal(2, VisioDocument.Load(path).Pages.Count);
        } finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void BookAddsReciprocalDeduplicatedPageLinksAndPreservesThemThroughSave() {
        var first = PlacementEnvelope(); var second = PlacementEnvelope();
        first.Title = second.Title = "Services";
        var links = new[] {
            new OfficeVisioVisualBookLink(1, "api", 2, "database", "api-database-a"),
            new OfficeVisioVisualBookLink(1, "api", 2, "database", "api-database-b")
        };

        OfficeVisioVisualBookResult book = new[] { first, second }.ToOfficeVisioBookWithNavigation(links);

        Assert.Equal(4, book.RequestedNavigationCount);
        Assert.Equal(2, book.CoalescedNavigationCount);
        Assert.Equal(0, book.OmittedNavigationCount);
        Assert.Equal(2, book.Navigations.Count);
        OfficeVisioVisualBookNavigationResult forward = Assert.Single(book.Navigations, navigation => !navigation.IsReturnLink);
        Assert.Equal(new[] { "api-database-a", "api-database-b" }, forward.RelationshipIds);
        VisioShape source = book.Pages[0].Page.Shapes.Single(shape => shape.Id == "api");
        VisioShape target = book.Pages[1].Page.Shapes.Single(shape => shape.Id == "database");
        VisioHyperlink sourceLink = Assert.Single(source.Hyperlinks);
        VisioHyperlink targetLink = Assert.Single(target.Hyperlinks);
        Assert.True(string.IsNullOrEmpty(sourceLink.Address));
        Assert.Equal("Services (2)", sourceLink.SubAddress);
        Assert.True(string.IsNullOrEmpty(targetLink.Address));
        Assert.Equal("Services", targetLink.SubAddress);
        Assert.Equal("Services (2)", source.GetShapeDataValue("CFX.BookLink.1.TargetPage"));
        Assert.Equal("database", source.GetShapeDataValue("CFX.BookLink.1.TargetEntityId"));
        Assert.Equal("api-database-a,api-database-b", source.GetShapeDataValue("CFX.BookLink.1.RelationshipIds"));

        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
        try {
            book.Document.Save(path);
            Assert.Empty(VisioValidator.Validate(path));
            VisioDocument loaded = VisioDocument.Load(path);
            VisioHyperlink loadedForward = Assert.Single(loaded.Pages[0].Shapes.Single(shape => shape.Id == "api").Hyperlinks);
            VisioHyperlink loadedReturn = Assert.Single(loaded.Pages[1].Shapes.Single(shape => shape.Id == "database").Hyperlinks);
            Assert.True(string.IsNullOrEmpty(loadedForward.Address));
            Assert.Equal("Services (2)", loadedForward.SubAddress);
            Assert.Equal("Services", loadedReturn.SubAddress);
        } finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void BookBoundsNavigationPerEntityAndReportsOmittedTargets() {
        var pages = Enumerable.Range(1, 4).Select(number => {
            var envelope = PlacementEnvelope();
            envelope.Title = "Page " + number;
            return envelope;
        }).ToArray();
        var links = new[] {
            new OfficeVisioVisualBookLink(1, "api", 2, "database", "to-2"),
            new OfficeVisioVisualBookLink(1, "api", 3, "database", "to-3"),
            new OfficeVisioVisualBookLink(1, "api", 4, "database", "to-4")
        };
        var bookOptions = new OfficeVisioVisualBookOptions {
            IncludeReturnLinks = false,
            MaximumNavigationLinksPerEntity = 2
        };

        OfficeVisioVisualBookResult book = pages.ToOfficeVisioBookWithNavigation(links, bookOptions);

        Assert.Equal(3, book.RequestedNavigationCount);
        Assert.Equal(2, book.Navigations.Count);
        Assert.Equal(1, book.OmittedNavigationCount);
        VisioShape source = book.Pages[0].Page.Shapes.Single(shape => shape.Id == "api");
        Assert.Equal(2, source.Hyperlinks.Count);
        Assert.Equal("1", source.GetShapeDataValue("CFX.BookLink.Omitted"));
        Assert.Equal(new[] { "Page 2", "Page 3" }, source.Hyperlinks.Select(link => link.SubAddress));
    }

    [Fact]
    public void BookBoundsRequestedLinksAndCoalescedRelationshipMetadata() {
        var pages = new[] { PlacementEnvelope(), PlacementEnvelope() };
        var links = Enumerable.Range(0, 3)
            .Select(index => new OfficeVisioVisualBookLink(1, "api", 2, "database", "edge-" + index))
            .ToArray();

        Assert.Throws<ArgumentException>(() => pages.ToOfficeVisioBookWithNavigation(links,
            new OfficeVisioVisualBookOptions { MaximumRequestedLinks = 2 }));
        Assert.Throws<ArgumentException>(() => pages.ToOfficeVisioBookWithNavigation(links,
            new OfficeVisioVisualBookOptions { MaximumRelationshipIdsPerNavigation = 2 }));
        Assert.Throws<ArgumentException>(() => pages.ToOfficeVisioBookWithNavigation(links,
            new OfficeVisioVisualBookOptions { MaximumRelationshipIdCharactersPerNavigation = 10 }));

        OfficeVisioVisualBookResult book = pages.ToOfficeVisioBookWithNavigation(links);
        Assert.Equal(new[] { "edge-0", "edge-1", "edge-2" },
            Assert.Single(book.Navigations, navigation => !navigation.IsReturnLink).RelationshipIds);
    }

    [Theory]
    [InlineData(0, nameof(OfficeVisioVisualBookOptions.MaximumRequestedLinks))]
    [InlineData(1, nameof(OfficeVisioVisualBookOptions.MaximumRelationshipIdsPerNavigation))]
    [InlineData(2, nameof(OfficeVisioVisualBookOptions.MaximumRelationshipIdCharactersPerNavigation))]
    public void BookRejectsInvalidLimitsWithTheirOwnParameterName(int invalidLimit, string expectedParameter) {
        var options = new OfficeVisioVisualBookOptions();
        switch (invalidLimit) {
            case 0: options.MaximumRequestedLinks = 0; break;
            case 1: options.MaximumRelationshipIdsPerNavigation = 0; break;
            default: options.MaximumRelationshipIdCharactersPerNavigation = 0; break;
        }

        var exception = Assert.Throws<ArgumentOutOfRangeException>(() =>
            new[] { PlacementEnvelope(), PlacementEnvelope() }.ToOfficeVisioBookWithNavigation(
                Array.Empty<OfficeVisioVisualBookLink>(), options));
        Assert.Equal(expectedParameter, exception.ParamName);
    }

    [Fact]
    public void BookResolvesOriginalChartForgeXIdsAfterInterchangeBoundsThem() {
        var first = PlacementEnvelope();
        first.Nodes[0].Id = "bounded-node-id";
        first.Nodes[0].Extensions["chartforgex.sourceId"] = "original-node-id";
        first.Edges[0].SourceId = "bounded-node-id";
        var second = PlacementEnvelope();

        OfficeVisioVisualBookResult book = new[] { first, second }.ToOfficeVisioBookWithNavigation(new[] {
            new OfficeVisioVisualBookLink(1, "original-node-id", 2, "database")
        }, options: new OfficeVisioVisualOptions { IncludeShapeData = false });

        VisioShape projected = book.Pages[0].Page.Shapes.Single(shape => shape.Id == "bounded-node-id");
        Assert.Equal(book.Pages[1].Page.Name, Assert.Single(projected.Hyperlinks).SubAddress);
    }

    [Fact]
    public void BookNormalizesAliasesBeforeDeduplicationAndBounds() {
        var first = PlacementEnvelope();
        first.Nodes[0].Id = "bounded-node-id";
        first.Nodes[0].Extensions["chartforgex.sourceId"] = "original-node-id";
        first.Edges[0].SourceId = "bounded-node-id";
        var second = PlacementEnvelope();
        var links = new[] {
            new OfficeVisioVisualBookLink(1, "bounded-node-id", 2, "database", "projected"),
            new OfficeVisioVisualBookLink(1, "original-node-id", 2, "database", "original")
        };

        OfficeVisioVisualBookResult book = new[] { first, second }.ToOfficeVisioBookWithNavigation(
            links,
            new OfficeVisioVisualBookOptions { IncludeReturnLinks = false, MaximumNavigationLinksPerEntity = 1 },
            new OfficeVisioVisualOptions { IncludeShapeData = false });

        Assert.Equal(2, book.RequestedNavigationCount);
        Assert.Equal(1, book.CoalescedNavigationCount);
        Assert.Equal(0, book.OmittedNavigationCount);
        Assert.Single(book.Navigations);
        Assert.Equal(new[] { "projected", "original" }, book.Navigations[0].RelationshipIds);
        Assert.Single(book.Pages[0].Page.Shapes.Single(shape => shape.Id == "bounded-node-id").Hyperlinks);
    }

    [Fact]
    public void BookRejectsAliasesThatIdentifyDifferentShapes() {
        var first = PlacementEnvelope();
        first.Nodes[0].Id = "bounded-node-id";
        first.Nodes[0].Extensions["chartforgex.sourceId"] = "original-node-id";
        first.Nodes[1].Id = "original-node-id";
        first.Edges[0].SourceId = "bounded-node-id";
        first.Edges[0].TargetId = "original-node-id";
        var second = PlacementEnvelope();

        ArgumentException exception = Assert.Throws<ArgumentException>(() =>
            new[] { first, second }.ToOfficeVisioBookWithNavigation(new[] {
                new OfficeVisioVisualBookLink(1, "original-node-id", 2, "database")
            }));

        Assert.Contains("ambiguous", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void BookRejectsOutOfRangePagesAndUnprojectedEntities() {
        var pages = new[] { PlacementEnvelope(), PlacementEnvelope() };
        Assert.Throws<ArgumentOutOfRangeException>(() => pages.ToOfficeVisioBookWithNavigation(new[] {
            new OfficeVisioVisualBookLink(1, "api", 3, "database")
        }));
        Assert.Throws<ArgumentException>(() => pages.ToOfficeVisioBookWithNavigation(new[] {
            new OfficeVisioVisualBookLink(1, "missing", 2, "database")
        }));
        Assert.Equal(2, pages.ToOfficeVisioBook((OfficeVisioVisualOptions?)null).Pages.Count);
    }

    [Fact]
    public void PreservedOverflowIsReportedAndCanBeRejected() {
        var envelope = PlacementEnvelope();
        envelope.Nodes[0].X = -10;
        var result = envelope.ToOfficeVisio();
        Assert.Contains(result.Report.Diagnostics, item => item.Code == OfficeVisioVisualDiagnosticCode.GeometryOutsidePage && item.EntityId == "api");
        var options = new OfficeVisioVisualOptions();
        options.RejectedDiagnostics.Add(OfficeVisioVisualDiagnosticCode.GeometryOutsidePage);
        Assert.Throws<OfficeVisioVisualFidelityException>(() => envelope.ToOfficeVisio(options));
    }

    [Fact]
    public void SelfRouteKeepsExplicitPortsAndLosslessPolicyRejectsNormalization() {
        var envelope = PlacementEnvelope();
        var edge = envelope.Edges[0];
        edge.TargetId = "api";
        edge.TargetPortId = "out";
        edge.Topology!.Waypoints.Clear();
        var result = envelope.ToOfficeVisio(new OfficeVisioVisualOptions { PixelsPerInch = 100 });
        var connector = Assert.Single(result.Page.Connectors);
        Assert.Equal(.75, connector.FromConnectionPoint!.Y, 6);
        Assert.Equal(.75, connector.ToConnectionPoint!.Y, 6);
        Assert.True(connector.Waypoints.Count >= 3);
        Assert.Throws<OfficeVisioVisualFidelityException>(() => envelope.ToOfficeVisio(new OfficeVisioVisualOptions { RequireLossless = true }));
    }

    [Fact]
    public void NativeGraphUsesSourceCardColorsAndAllowsExplicitThemeOverride() {
        var chart = TopologyChart.Create().WithTheme(TopologyTheme.Light()).AddNode("a", "Service", 100, 100);
        chart.Theme!.Card = "#F0F4FA";
        chart.Theme.Foreground = "#14243A";
        var envelope = chart.ToVisualArtifact().ToInterchangeEnvelope();
        var result = envelope.ToOfficeVisio();
        var shape = result.Page.Shapes.Single(item => item.Id == "a");
        Assert.Equal(OfficeIMO.Drawing.OfficeColor.FromRgb(240, 244, 250), shape.FillColor);
        Assert.Equal(OfficeIMO.Drawing.OfficeColor.FromRgb(20, 36, 58), shape.TextStyle!.Color);
        Assert.Equal("Arial", shape.TextStyle.FontFamily);
        var custom = VisioStyleTheme.Technical();
        var overridden = envelope.ToOfficeVisio(new OfficeVisioVisualOptions { NativeTheme = custom });
        Assert.Equal(custom.Primary.FillColor, overridden.Page.Shapes.Single(item => item.Id == "a").FillColor);
    }

    private static VisualArtifactInterchangeEnvelope PlacementEnvelope() {
        var envelope = TopologyEnvelope("placed");
        envelope.Width = 1000; envelope.Height = 800;
        var api = TopologyNode("api", "API");
        api.X = 100; api.Y = 100; api.Width = 200; api.Height = 100;
        api.Ports.Add(new VisualArtifactInterchangePort { Id = "out", Side = TopologyEdgePort.Right, Offset = .25 });
        var database = TopologyNode("database", "Database");
        database.X = 600; database.Y = 400; database.Width = 200; database.Height = 100;
        envelope.Nodes.Add(api); envelope.Nodes.Add(database);
        var edge = TopologyEdge("link", VisualLinkDirection.Forward);
        edge.SourcePortId = "out";
        edge.Topology!.Routing = TopologyEdgeRouting.Straight;
        edge.Topology.Waypoints.Add(new VisualArtifactInterchangePoint { X = 400, Y = 125 });
        edge.Topology.Waypoints.Add(new VisualArtifactInterchangePoint { X = 400, Y = 450 });
        envelope.Edges.Add(edge);
        return envelope;
    }
}
