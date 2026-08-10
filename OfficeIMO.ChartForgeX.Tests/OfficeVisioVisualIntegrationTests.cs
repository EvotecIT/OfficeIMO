using System;
using System.IO;
using System.Linq;
using global::ChartForgeX.VisualArtifacts;
using OfficeIMO.ChartForgeX;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.ChartForgeX.Tests;

public sealed class OfficeVisioVisualIntegrationTests {
    [Fact]
    public void TopologyEnvelopeCreatesEditableValidatedVsdxWithShapeDataAndLinks() {
        VisualArtifactInterchangeEnvelope envelope = CreateTopologyEnvelope();

        OfficeVisioVisualConversionResult result = envelope.ToUtf8Json().ToOfficeVisio(
            new OfficeVisioVisualOptions { PageName = "Service topology" });

        Assert.True(result.Report.IsNativeEditable);
        Assert.Equal("VisioGraphDiagramBuilder", result.Report.Projection);
        Assert.Equal(2, result.Report.NodeCount);
        Assert.Equal(1, result.Report.EdgeCount);
        Assert.Equal(1, result.Report.GroupCount);
        Assert.Equal(0, result.Report.AnnotationCount);
        Assert.True(result.Page.Width < 9D);
        Assert.Contains(result.Page.Shapes, shape => shape.Id == "api");
        Assert.Contains(result.Page.Shapes, shape => shape.Id == "data-zone");
        Assert.Contains(result.Page.Connectors, connector => connector.Id == "api-db");

        var api = result.Page.Shapes.Single(shape => shape.Id == "api");
        Assert.Equal("TopologyNode", api.GetShapeDataValue("CFX.Kind"));
        Assert.Equal("Platform", api.GetShapeDataValue("Metadata.Owner"));
        Assert.Equal("443", api.GetShapeDataValue("Detail.1.Port"));
        Assert.Contains(api.Hyperlinks, link => link.Address == "https://example.test/api");

        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
        try {
            result.Document.Save(path);
            Assert.Empty(VisioValidator.Validate(path));
            VisioDocument loaded = VisioDocument.Load(path);
            Assert.Equal("Platform", loaded.Pages[0].Shapes.Single(shape => shape.Id == "api").GetShapeDataValue("Metadata.Owner"));
            Assert.Contains(loaded.Pages[0].Connectors, connector => connector.Id == "api-db");
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void SequenceEnvelopeCreatesNativeParticipantsMessagesActivationsNotesAndFragments() {
        var envelope = new VisualArtifactInterchangeEnvelope {
            Id = "checkout",
            Kind = VisualArtifactKind.Sequence,
            Title = "Checkout sequence",
            Width = 720,
            Height = 480
        };
        envelope.Nodes.Add(Participant("customer", "Customer", "Actor", 0));
        envelope.Nodes.Add(Participant("api", "Orders API", "Control", 1));
        envelope.Edges.Add(Message("request", "customer", "api", "Create order", 0, activates: true));
        envelope.Edges.Add(Message("response", "api", "customer", "Created", 1, deactivates: true, dashed: true));
        var note = new VisualArtifactInterchangeAnnotation {
            Id = "retry-note",
            Kind = "SequenceNote",
            Text = "Retry window",
            Placement = "RightOf",
            StartIndex = 0,
            EndIndex = 0
        };
        note.TargetIds.Add("api");
        envelope.Annotations.Add(note);
        envelope.Annotations.Add(new VisualArtifactInterchangeAnnotation {
            Id = "alt-block",
            Kind = "SequenceBlock:Alt",
            Text = "order accepted",
            StartIndex = 0,
            EndIndex = 1
        });

        OfficeVisioVisualConversionResult result = envelope.ToOfficeVisio();

        Assert.Equal("VisioSequenceDiagramBuilder", result.Report.Projection);
        Assert.Contains(result.Page.Shapes, shape => shape.Id == "customer");
        Assert.Contains(result.Page.Shapes, shape => shape.Id == "api");
        Assert.Contains(result.Page.Shapes, shape => shape.Id == "activation-1");
        Assert.Contains(result.Page.Shapes, shape => shape.Id == "retry-note");
        Assert.Contains(result.Page.Shapes, shape => shape.Id == "alt-block");
        Assert.Contains(result.Page.Connectors, connector => connector.Id == "request");
        Assert.Contains(result.Page.Connectors, connector => connector.Id == "response" && connector.LinePattern == 2);
        Assert.Equal(2, result.Report.AnnotationCount);
    }

    [Fact]
    public void UnsupportedArtifactKindFailsClosedInsteadOfClaimingEditableFidelity() {
        var envelope = new VisualArtifactInterchangeEnvelope {
            Id = "sales",
            Kind = VisualArtifactKind.Chart,
            Title = "Sales"
        };

        NotSupportedException exception = Assert.Throws<NotSupportedException>(() => envelope.ToOfficeVisio());
        Assert.Contains("separately rendered SVG", exception.Message);
    }

    [Fact]
    public void VisioOptionsRejectInvalidPageAndPixelDensity() {
        Assert.Throws<ArgumentException>(() => new OfficeVisioVisualOptions { PageName = " " });
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeVisioVisualOptions { PixelsPerInch = 0D });
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeVisioVisualOptions { PixelsPerInch = double.NaN });
    }

    [Fact]
    public void NaturalPageSizingIsExplicitInsteadOfAddingDefaultCanvasWhitespace() {
        VisualArtifactInterchangeEnvelope envelope = CreateTopologyEnvelope();

        OfficeVisioVisualConversionResult fitted = envelope.ToOfficeVisio();
        OfficeVisioVisualConversionResult natural = envelope.ToOfficeVisio(new OfficeVisioVisualOptions {
            UseNaturalPageSize = true,
            PixelsPerInch = 100D
        });

        Assert.True(fitted.Page.Width < natural.Page.Width);
        Assert.True(natural.Page.Width >= 9D);
    }

    [Fact]
    public void FidelityCountsOnlyNativeObjectsAndWarnsForUnmappedGraphAnnotations() {
        VisualArtifactInterchangeEnvelope envelope = CreateTopologyEnvelope();
        envelope.Annotations.Add(new VisualArtifactInterchangeAnnotation {
            Id = "graph-note",
            Kind = "Note",
            Text = "Retained only in the semantic envelope"
        });

        OfficeVisioVisualConversionResult result = envelope.ToOfficeVisio(new OfficeVisioVisualOptions { IncludeGroups = false });

        Assert.Equal(0, result.Report.GroupCount);
        Assert.Equal(0, result.Report.AnnotationCount);
        Assert.Contains(result.Report.Warnings, warning => warning.Contains("graph-note"));
    }

    private static VisualArtifactInterchangeEnvelope CreateTopologyEnvelope() {
        var envelope = new VisualArtifactInterchangeEnvelope {
            Id = "service-topology",
            Kind = VisualArtifactKind.Topology,
            Title = "Service topology",
            Layout = "Layered",
            Direction = "LeftToRight",
            Width = 900,
            Height = 520
        };
        envelope.Groups.Add(new VisualArtifactInterchangeGroup { Id = "data-zone", Kind = "TopologyGroup", Label = "Data zone" });
        var api = new VisualArtifactInterchangeNode {
            Id = "api",
            Kind = "TopologyNode",
            Label = "API",
            Status = "Healthy",
            Href = "https://example.test/api",
            Tooltip = "API runbook"
        };
        api.Metadata["Owner"] = "Platform";
        api.Details.Add(new VisualArtifactInterchangeDetail { Label = "Port", Value = "443" });
        api.Ports.Add(new VisualArtifactInterchangePort { Id = "out", Side = "Right", Offset = 0.5D });
        envelope.Nodes.Add(api);
        envelope.Nodes.Add(new VisualArtifactInterchangeNode {
            Id = "database",
            Kind = "Database",
            Label = "Database",
            GroupId = "data-zone"
        });
        envelope.Edges.Add(new VisualArtifactInterchangeEdge {
            Id = "api-db",
            Kind = "Data",
            SourceId = "api",
            TargetId = "database",
            Label = "queries",
            Direction = "Forward",
            SourcePortId = "out"
        });
        return envelope;
    }

    private static VisualArtifactInterchangeNode Participant(string id, string label, string kind, int order) {
        var participant = new VisualArtifactInterchangeNode { Id = id, Label = label, Kind = kind };
        participant.Metadata["sequence.order"] = order.ToString();
        participant.Metadata["sequence.implicit"] = "false";
        return participant;
    }

    private static VisualArtifactInterchangeEdge Message(
        string id,
        string source,
        string target,
        string label,
        int order,
        bool activates = false,
        bool deactivates = false,
        bool dashed = false) {
        var message = new VisualArtifactInterchangeEdge {
            Id = id,
            Kind = "SequenceMessage",
            SourceId = source,
            TargetId = target,
            Label = label,
            Order = order,
            LineStyle = dashed ? "Dashed" : "Solid"
        };
        message.Metadata["sequence.activatesTarget"] = activates ? "true" : "false";
        message.Metadata["sequence.deactivates"] = deactivates ? "true" : "false";
        return message;
    }
}
