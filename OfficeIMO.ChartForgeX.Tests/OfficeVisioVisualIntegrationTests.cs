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
        Assert.Equal("Secondary", api.GetShapeDataValue("Metadata.owner [2]"));
        Assert.Contains(result.Report.Warnings, warning => warning.Contains("case-insensitive"));
        Assert.Equal("443", api.GetShapeDataValue("Detail.1.Port"));
        Assert.Equal("health", api.GetShapeDataValue("Detail.1.Icon"));
        Assert.Equal("Healthy", api.GetShapeDataValue("Detail.1.Status"));
        Assert.Equal("#22AA66", api.GetShapeDataValue("Detail.1.Color"));
        Assert.Equal("TCP", api.GetShapeDataValue("Detail.1.Metadata.Protocol"));
        Assert.Equal("egress", api.GetShapeDataValue("Port.1.Label"));
        Assert.Equal("primary", api.GetShapeDataValue("Port.1.Metadata.Role"));
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
        VisualArtifactInterchangeNode apiParticipant = Participant("api", "Orders API", "Control", 1);
        apiParticipant.Subtitle = "v2";
        apiParticipant.Status = "Healthy";
        apiParticipant.Href = "https://example.test/orders";
        apiParticipant.Tooltip = "Orders runbook";
        apiParticipant.Metadata["Owner"] = "Commerce";
        apiParticipant.Details.Add(new VisualArtifactInterchangeDetail { Label = "Region", Value = "EU" });
        envelope.Nodes.Add(apiParticipant);
        envelope.Nodes.Add(Participant("activation-1", "Reserved activation id", "Participant", 2));
        VisualArtifactInterchangeEdge request = Message("request", "customer", "api", "Create order", 0, activates: true);
        request.SecondaryLabel = "async";
        request.TertiaryLabel = "audited";
        request.SourceLabel = "client";
        request.TargetLabel = "service";
        request.Status = "Healthy";
        request.Href = "https://example.test/create-order";
        request.Tooltip = "Create order contract";
        request.Metadata["Owner"] = "Commerce";
        envelope.Edges.Add(request);
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
        Assert.Contains(result.Page.Shapes, shape => shape.Id == "activation-2");
        Assert.Contains(result.Page.Shapes, shape => shape.Id == "retry-note");
        Assert.Contains(result.Page.Shapes, shape => shape.Id == "alt-block");
        Assert.Contains(result.Page.Connectors, connector => connector.Id == "request");
        Assert.Contains(result.Page.Connectors, connector => connector.Id == "response" && connector.LinePattern == 2);
        Assert.Equal(2, result.Report.AnnotationCount);
        VisioShape api = result.Page.Shapes.Single(shape => shape.Id == "api");
        VisioShape retryNote = result.Page.Shapes.Single(shape => shape.Id == "retry-note");
        VisioConnector requestConnector = result.Page.Connectors.Single(connector => connector.Id == "request");
        Assert.Equal("Healthy", api.GetShapeDataValue("CFX.Status"));
        Assert.Equal("Commerce", api.GetShapeDataValue("Metadata.Owner"));
        Assert.Equal("EU", api.GetShapeDataValue("Detail.1.Region"));
        Assert.Equal("Orders API" + Environment.NewLine + "v2", api.Text);
        Assert.Contains(api.Hyperlinks, link => link.Address == "https://example.test/orders" && link.Description == "Orders runbook");
        Assert.Equal("request", requestConnector.GetShapeDataValue("CFX.Id"));
        Assert.Equal("Healthy", requestConnector.GetShapeDataValue("CFX.Status"));
        Assert.Equal("Commerce", requestConnector.GetShapeDataValue("Metadata.Owner"));
        Assert.Equal("Create order | async | audited", requestConnector.Label);
        Assert.Equal("client", requestConnector.GetShapeDataValue("CFX.SourceLabel"));
        Assert.Equal("service", requestConnector.GetShapeDataValue("CFX.TargetLabel"));
        Assert.Equal("0", requestConnector.GetShapeDataValue("CFX.Order"));
        Assert.Contains(requestConnector.Hyperlinks, link => link.Address == "https://example.test/create-order" && link.Description == "Create order contract");
        Assert.True(retryNote.PinX - retryNote.Width / 2D > api.PinX + api.Width / 2D);
    }

    [Fact]
    public void SequenceEqualOrdersPreserveEnvelopeCollectionOrder() {
        var envelope = new VisualArtifactInterchangeEnvelope { Id = "stable-order", Kind = VisualArtifactKind.Sequence };
        envelope.Nodes.Add(new VisualArtifactInterchangeNode { Id = "z-first", Label = "First" });
        envelope.Nodes.Add(new VisualArtifactInterchangeNode { Id = "a-second", Label = "Second" });
        envelope.Edges.Add(Message("z-first-message", "z-first", "a-second", "First message", 0));
        envelope.Edges.Add(Message("a-second-message", "a-second", "z-first", "Second message", 0));

        OfficeVisioVisualConversionResult result = envelope.ToOfficeVisio();

        Assert.True(result.Page.Shapes.Single(shape => shape.Id == "z-first").PinX < result.Page.Shapes.Single(shape => shape.Id == "a-second").PinX);
        Assert.Equal(new[] { "z-first-message", "a-second-message" }, result.Page.Connectors
            .Where(connector => connector.Id == "z-first-message" || connector.Id == "a-second-message")
            .Select(connector => connector.Id)
            .ToArray());
    }

    [Fact]
    public void SequencePreservesLateAnnotationRowsAndDeterministicOpenActivations() {
        var envelope = new VisualArtifactInterchangeEnvelope { Id = "late-rows", Kind = VisualArtifactKind.Sequence };
        envelope.Nodes.Add(Participant("caller", "Caller", "Actor", 0));
        envelope.Nodes.Add(Participant("z-worker", "Z worker", "Control", 1));
        envelope.Nodes.Add(Participant("a-worker", "A worker", "Control", 2));
        envelope.Edges.Add(Message("activate-z", "caller", "z-worker", "Z", 0, activates: true));
        envelope.Edges.Add(Message("activate-a", "caller", "a-worker", "A", 1, activates: true));
        var note = new VisualArtifactInterchangeAnnotation { Id = "late-note", Kind = "SequenceNote", Text = "Later", StartIndex = 5, EndIndex = 5 };
        note.TargetIds.Add("caller");
        envelope.Annotations.Add(note);
        envelope.Annotations.Add(new VisualArtifactInterchangeAnnotation { Id = "late-block", Kind = "SequenceBlock:Opt", Text = "Later block", StartIndex = 4, EndIndex = 6 });

        OfficeVisioVisualConversionResult result = envelope.ToOfficeVisio();

        Assert.Equal("5", result.Page.Shapes.Single(shape => shape.Id == "late-note").GetUserCellValue("OfficeIMO.SequenceRowIndex"));
        Assert.Equal("4", result.Page.Shapes.Single(shape => shape.Id == "late-block").GetUserCellValue("OfficeIMO.SequenceStartRowIndex"));
        Assert.Equal("6", result.Page.Shapes.Single(shape => shape.Id == "late-block").GetUserCellValue("OfficeIMO.SequenceEndRowIndex"));
        Assert.Equal("a-worker", result.Page.Shapes.Single(shape => shape.Id == "activation-1").GetUserCellValue("OfficeIMO.SequenceParticipantId"));
        Assert.Equal("z-worker", result.Page.Shapes.Single(shape => shape.Id == "activation-2").GetUserCellValue("OfficeIMO.SequenceParticipantId"));
    }

    [Fact]
    public void SequenceUnknownNumericParticipantKindFallsBackAndReportsFidelity() {
        var envelope = new VisualArtifactInterchangeEnvelope { Id = "unknown-kind", Kind = VisualArtifactKind.Sequence };
        envelope.Nodes.Add(new VisualArtifactInterchangeNode { Id = "participant", Label = "Participant", Kind = "999" });

        OfficeVisioVisualConversionResult result = envelope.ToOfficeVisio();

        Assert.Contains(result.Page.Shapes, shape => shape.Id == "participant");
        Assert.Contains(result.Report.Warnings, warning => warning.Contains("kind '999'") && warning.Contains("generic participant"));
    }

    [Fact]
    public void ParticipantOnlySequenceProjectsNotesAndFragmentsAtRowZero() {
        var envelope = new VisualArtifactInterchangeEnvelope { Id = "participant-only", Kind = VisualArtifactKind.Sequence };
        envelope.Nodes.Add(Participant("service", "Service", "Control", 0));
        var note = new VisualArtifactInterchangeAnnotation { Id = "note", Kind = "SequenceNote", Text = "Ready", StartIndex = 0, EndIndex = 0 };
        note.TargetIds.Add("service");
        envelope.Annotations.Add(note);
        var fragment = new VisualArtifactInterchangeAnnotation { Id = "fragment", Kind = "SequenceBlock:Opt", Text = "cached", StartIndex = 0, EndIndex = 0 };
        fragment.TargetIds.Add("service");
        envelope.Annotations.Add(fragment);

        OfficeVisioVisualConversionResult result = envelope.ToOfficeVisio();

        Assert.Equal(2, result.Report.AnnotationCount);
        Assert.Contains(result.Page.Shapes, shape => shape.Id == "note");
        Assert.Contains(result.Page.Shapes, shape => shape.Id == "fragment");
    }

    [Fact]
    public void SequencePageUsesEnvelopeDimensionsOnlyWhenNaturalSizingIsRequested() {
        var envelope = new VisualArtifactInterchangeEnvelope { Id = "wide-sequence", Kind = VisualArtifactKind.Sequence, Width = 2400, Height = 1200 };
        envelope.Nodes.Add(Participant("caller", "Caller", "Actor", 0));
        envelope.Nodes.Add(Participant("service", "Service", "Control", 1));
        envelope.Edges.Add(Message("call", "caller", "service", "Call", 0));

        OfficeVisioVisualConversionResult fitted = envelope.ToOfficeVisio();
        OfficeVisioVisualConversionResult natural = envelope.ToOfficeVisio(new OfficeVisioVisualOptions { UseNaturalPageSize = true, PixelsPerInch = 100D });

        Assert.True(fitted.Page.Width < natural.Page.Width);
        Assert.True(fitted.Page.Width < 9D);
        Assert.True(natural.Page.Width >= 24D);
    }

    [Fact]
    public void SequenceSemanticIdsAreMappedAwayFromGeneratedVisioHelpers() {
        var envelope = new VisualArtifactInterchangeEnvelope { Id = "collisions", Kind = VisualArtifactKind.Sequence, Title = "Collisions" };
        envelope.Nodes.Add(Participant("api", "API", "Control", 0));
        envelope.Nodes.Add(Participant("api-lifeline", "Worker", "Participant", 1));
        envelope.Edges.Add(Message("api-lifeline-end", "api", "api-lifeline", "Dispatch", 0, activates: true));
        var fragment = new VisualArtifactInterchangeAnnotation {
            Id = "message-api-lifeline-end-from",
            Kind = "SequenceBlock:Opt",
            Text = "mapped",
            StartIndex = 0,
            EndIndex = 0
        };
        fragment.TargetIds.Add("api-lifeline");
        envelope.Annotations.Add(fragment);

        OfficeVisioVisualConversionResult result = envelope.ToOfficeVisio();

        VisioShape api = result.Page.Shapes.Single(shape => shape.Id == "api");
        VisioShape worker = result.Page.Shapes.Single(shape => shape.GetShapeDataValue("CFX.Id") == "api-lifeline");
        VisioConnector message = result.Page.Connectors.Single(connector => connector.GetShapeDataValue("CFX.Id") == "api-lifeline-end");
        VisioShape nativeFragment = result.Page.Shapes.Single(shape => shape.GetShapeDataValue("CFX.Id") == "message-api-lifeline-end-from");
        Assert.NotEqual("api-lifeline", worker.Id);
        Assert.NotEqual("api-lifeline-end", message.Id);
        Assert.NotEqual("message-api-lifeline-end-from", nativeFragment.Id);
        Assert.Equal(message.Id + "-from", message.From.Id);
        Assert.Equal(message.Id + "-to", message.To.Id);
        Assert.Contains(result.Report.Warnings, warning => warning.Contains("native Visio helper shapes"));
        Assert.Equal("api", api.GetShapeDataValue("CFX.Id"));
        Assert.Equal("api-lifeline", result.Envelope.Nodes.Single(node => node.Label == "Worker").Id);

        OfficeVisioVisualConversionResult withoutShapeData = envelope.ToOfficeVisio(new OfficeVisioVisualOptions { IncludeShapeData = false });
        Assert.Equal("api-lifeline", withoutShapeData.Page.Shapes.Single(shape => shape.Id != "api" && shape.Data.ContainsKey("CFX.Id") && shape.Data["CFX.Id"] == "api-lifeline").Data["CFX.Id"]);
        Assert.Equal("api-lifeline-end", withoutShapeData.Page.Connectors.Single(connector => connector.Data.TryGetValue("CFX.Id", out string? id) && id == "api-lifeline-end").Data["CFX.Id"]);
        Assert.Equal("message-api-lifeline-end-from", withoutShapeData.Page.Shapes.Single(shape => shape.Data.TryGetValue("CFX.Id", out string? id) && id == "message-api-lifeline-end-from").Data["CFX.Id"]);
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
        try {
            withoutShapeData.Document.Save(path);
            VisioDocument loaded = VisioDocument.Load(path);
            Assert.Contains(loaded.Pages[0].Shapes, shape => shape.Data.TryGetValue("CFX.Id", out string? id) && id == "api-lifeline");
            Assert.Contains(loaded.Pages[0].Connectors, connector => connector.Data.TryGetValue("CFX.Id", out string? id) && id == "api-lifeline-end");
            Assert.Contains(loaded.Pages[0].Shapes, shape => shape.Data.TryGetValue("CFX.Id", out string? id) && id == "message-api-lifeline-end-from");
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void GraphProjectionPreservesForwardBackwardAndBidirectionalArrows() {
        VisualArtifactInterchangeEnvelope envelope = CreateTopologyEnvelope();
        envelope.Edges.Clear();
        envelope.Edges.Add(new VisualArtifactInterchangeEdge { Id = "forward", SourceId = "api", TargetId = "database", Direction = "Forward" });
        envelope.Edges.Add(new VisualArtifactInterchangeEdge { Id = "backward", SourceId = "api", TargetId = "database", Direction = "Backward" });
        envelope.Edges.Add(new VisualArtifactInterchangeEdge { Id = "both", SourceId = "api", TargetId = "database", Direction = "Bidirectional" });

        OfficeVisioVisualConversionResult result = envelope.ToOfficeVisio();

        VisioConnector forward = result.Page.Connectors.Single(connector => connector.Id == "forward");
        VisioConnector backward = result.Page.Connectors.Single(connector => connector.Id == "backward");
        VisioConnector both = result.Page.Connectors.Single(connector => connector.Id == "both");
        Assert.Equal(EndArrow.None, forward.BeginArrow);
        Assert.Equal(EndArrow.Triangle, forward.EndArrow);
        Assert.Equal(EndArrow.Triangle, backward.BeginArrow);
        Assert.Equal(EndArrow.None, backward.EndArrow);
        Assert.Equal(EndArrow.Triangle, both.BeginArrow);
        Assert.Equal(EndArrow.Triangle, both.EndArrow);
    }

    [Fact]
    public void TopologyLayoutTokensMapToNativeStrategiesAndReportUnsupportedLayouts() {
        var matrix = new VisualArtifactInterchangeEnvelope { Id = "matrix", Kind = VisualArtifactKind.Topology, Layout = "Matrix" };
        for (int index = 0; index < 4; index++) matrix.Nodes.Add(new VisualArtifactInterchangeNode { Id = "node-" + index, Label = "Node " + index });

        OfficeVisioVisualConversionResult matrixResult = matrix.ToOfficeVisio();
        Assert.True(matrixResult.Page.Shapes.Where(shape => shape.Id.StartsWith("node-", StringComparison.Ordinal)).Select(shape => shape.PinX).Distinct().Count() > 1);
        Assert.True(matrixResult.Page.Shapes.Where(shape => shape.Id.StartsWith("node-", StringComparison.Ordinal)).Select(shape => shape.PinY).Distinct().Count() > 1);

        matrix.Layout = "Geographic";
        OfficeVisioVisualConversionResult geographicResult = matrix.ToOfficeVisio();
        Assert.Contains(geographicResult.Report.Warnings, warning => warning.Contains("Geographic") && warning.Contains("layered layout"));
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
        envelope.Metadata["Owner"] = "Platform";
        envelope.Annotations.Add(new VisualArtifactInterchangeAnnotation {
            Id = "graph-note",
            Kind = "Note",
            Text = "Retained only in the semantic envelope"
        });

        OfficeVisioVisualConversionResult result = envelope.ToOfficeVisio(new OfficeVisioVisualOptions { IncludeGroups = false });

        Assert.Equal(0, result.Report.GroupCount);
        Assert.Equal(0, result.Report.AnnotationCount);
        Assert.Contains(result.Report.Warnings, warning => warning.Contains("graph-note"));
        Assert.Contains(result.Report.Warnings, warning => warning.Contains("Artifact-level metadata"));
        Assert.Contains(result.Report.Warnings, warning => warning.Contains("native Visio graph layout selected connector sides"));
        VisioConnector edge = result.Page.Connectors.Single(connector => connector.Id == "api-db");
        Assert.Equal("out", edge.GetShapeDataValue("CFX.SourcePortId"));
        Assert.Equal("0", edge.GetShapeDataValue("CFX.Order"));
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
        api.Metadata["owner"] = "Secondary";
        var detail = new VisualArtifactInterchangeDetail { Label = "Port", Value = "443", IconId = "health", Status = "Healthy", Color = "#22AA66" };
        detail.Metadata["Protocol"] = "TCP";
        api.Details.Add(detail);
        var port = new VisualArtifactInterchangePort { Id = "out", Side = "Right", Offset = 0.5D, Label = "egress" };
        port.Metadata["Role"] = "primary";
        api.Ports.Add(port);
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
