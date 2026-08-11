using System;
using System.IO;
using System.Linq;
using global::ChartForgeX.Primitives;
using global::ChartForgeX.Topology;
using global::ChartForgeX.VisualArtifacts;
using OfficeIMO.ChartForgeX;
using OfficeIMO.Visio;
using Xunit;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.ChartForgeX.Tests;

public sealed class OfficeVisioVisualIntegrationTests {
    [Fact]
    public void TopologyEnvelopeCreatesEditableValidatedVsdxWithShapeDataAndLinks() {
        VisualArtifactInterchangeEnvelope envelope = CreateTopologyEnvelope();

        OfficeVisioVisualConversionResult result = envelope.ToUtf8Json().ToOfficeVisio(
            new OfficeVisioVisualOptions { PageName = "Service topology" });

        Assert.True(result.Report.AllProjectedObjectsEditable);
        Assert.Equal(OfficeVisioVisualProjectionKind.Graph, result.Report.Projection);
        Assert.Equal(2, result.Report.NodeCount);
        Assert.Equal(1, result.Report.EdgeCount);
        Assert.Equal(1, result.Report.GroupCount);
        Assert.Equal(0, result.Report.AnnotationCount);
        Assert.True(result.Page.Width < 9D);
        Assert.Contains(result.Page.Shapes, shape => shape.Id == "api");
        Assert.Contains(result.Page.Shapes, shape => shape.Id == "data-zone");
        Assert.Contains(result.Page.Connectors, connector => connector.Id == "api-db");

        var api = result.Page.Shapes.Single(shape => shape.Id == "api");
        Assert.Equal("Service", api.GetShapeDataValue("CFX.Kind"));
        Assert.Equal("Platform", api.GetShapeDataValue("Extension.Owner"));
        Assert.Equal("Secondary", api.GetShapeDataValue("Extension.owner [2]"));
        Assert.Equal("caller-owned", api.GetShapeDataValue("Extension.Metric.latency"));
        Assert.Equal("42 ms", api.GetShapeDataValue("Metric.latency"));
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == OfficeVisioVisualDiagnosticCode.ExtensionKeyRenamed &&
            diagnostic.Severity == OfficeVisioVisualDiagnosticSeverity.Information);
        Assert.Equal("443", api.GetShapeDataValue("Detail.1.Port"));
        Assert.Equal("health", api.GetShapeDataValue("Detail.1.Icon"));
        Assert.Equal("Healthy", api.GetShapeDataValue("Detail.1.Status"));
        Assert.Equal("#22AA66", api.GetShapeDataValue("Detail.1.Color"));
        Assert.Equal("TCP", api.GetShapeDataValue("Detail.1.Extension.Protocol"));
        Assert.Equal("egress", api.GetShapeDataValue("Port.1.Label"));
        Assert.Equal("primary", api.GetShapeDataValue("Port.1.Extension.Role"));
        Assert.Contains(api.Hyperlinks, link => link.Address == "https://example.test/api");

        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
        try {
            result.Document.Save(path);
            Assert.Empty(VisioValidator.Validate(path));
            VisioDocument loaded = VisioDocument.Load(path);
            Assert.Equal("Platform", loaded.Pages[0].Shapes.Single(shape => shape.Id == "api").GetShapeDataValue("Extension.Owner"));
            Assert.Contains(loaded.Pages[0].Connectors, connector => connector.Id == "api-db");
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DetailLabelsCannotOverwriteReservedShapeDataFields() {
        VisualArtifactInterchangeEnvelope envelope = CreateTopologyEnvelope();
        envelope.Nodes.Single(node => node.Id == "api").Details.Add(new VisualArtifactInterchangeDetail {
            Label = "Icon",
            Value = "Redis"
        });

        OfficeVisioVisualConversionResult result = envelope.ToOfficeVisio();

        VisioShape api = result.Page.Shapes.Single(shape => shape.Id == "api");
        Assert.Equal("Redis", api.GetShapeDataValue("Detail.2.Field.Icon"));
        Assert.Null(api.GetShapeDataValue("Detail.2.Icon"));
        Assert.Equal("Icon", api.GetShapeDataValue("Detail.2.Label"));
        Assert.Equal("Redis", api.GetShapeDataValue("Detail.2.Value"));
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == OfficeVisioVisualDiagnosticCode.DetailFieldRenamed &&
            diagnostic.Severity == OfficeVisioVisualDiagnosticSeverity.Information &&
            diagnostic.Message.Contains("Detail.2.Field.Icon"));
    }

    [Fact]
    public void LosslessShapeDataRenamesRemainInformational() {
        VisualArtifactInterchangeEnvelope envelope = TopologyEnvelope("informational-renames");
        VisualArtifactInterchangeNode node = TopologyNode("service", "Service", TopologyNodeKind.Service);
        node.Extensions["Owner"] = "Platform";
        node.Extensions["owner"] = "Secondary";
        envelope.Nodes.Add(node);

        OfficeVisioVisualConversionResult result = envelope.ToOfficeVisio();

        Assert.False(result.Report.HasSemanticLoss);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == OfficeVisioVisualDiagnosticCode.ExtensionKeyRenamed &&
            diagnostic.Severity == OfficeVisioVisualDiagnosticSeverity.Information);
        Assert.Equal("Platform", result.Page.Shapes.Single(shape => shape.Id == "service").GetShapeDataValue("Extension.Owner"));
        Assert.Equal("Secondary", result.Page.Shapes.Single(shape => shape.Id == "service").GetShapeDataValue("Extension.owner [2]"));
    }

    [Fact]
    public void SequenceEnvelopeCreatesNativeParticipantsMessagesActivationsNotesAndFragments() {
        VisualArtifactInterchangeEnvelope envelope = SequenceEnvelope("checkout", "Checkout sequence", 720, 480);
        envelope.Kind = VisualArtifactKind.Mermaid;
        VisualArtifactInterchangeNode customer = Participant("customer", "Customer", SequenceArtifactParticipantKind.Actor, 0);
        customer.Tooltip = "Checkout user";
        envelope.Nodes.Add(customer);
        VisualArtifactInterchangeNode apiParticipant = Participant("api", "Orders API", SequenceArtifactParticipantKind.Control, 1);
        apiParticipant.Subtitle = "v2";
        apiParticipant.Status = "Healthy";
        apiParticipant.Href = "https://example.test/orders";
        apiParticipant.Tooltip = "Orders runbook";
        apiParticipant.Extensions["Owner"] = "Commerce";
        apiParticipant.Details.Add(new VisualArtifactInterchangeDetail { Label = "Region", Value = "EU" });
        envelope.Nodes.Add(apiParticipant);
        envelope.Nodes.Add(Participant("activation-1", "Reserved activation id", SequenceArtifactParticipantKind.Participant, 2));
        VisualArtifactInterchangeEdge request = Message("request", "customer", "api", "Create order", 0, activates: true, kind: SequenceArtifactMessageKind.Async);
        request.SecondaryLabel = "async";
        request.TertiaryLabel = "audited";
        request.SourceLabel = "client";
        request.TargetLabel = "service";
        request.Status = "Healthy";
        request.Href = "https://example.test/create-order";
        request.Tooltip = "Create order contract";
        request.Extensions["Owner"] = "Commerce";
        envelope.Edges.Add(request);
        VisualArtifactInterchangeEdge response = Message("response", "api", "customer", "Created", 1, deactivates: true, dashed: true, kind: SequenceArtifactMessageKind.Return);
        response.Tooltip = "Created response";
        envelope.Edges.Add(response);
        var note = new VisualArtifactInterchangeAnnotation {
            Id = "retry-note",
            Role = VisualArtifactInterchangeAnnotationRole.SequenceNote,
            Kind = "SequenceNote",
            Text = "Retry window",
            Placement = "RightOf",
            StartIndex = 0,
            EndIndex = 0,
            Sequence = new VisualArtifactInterchangeSequenceAnnotation { NotePlacement = SequenceArtifactNotePlacement.RightOf }
        };
        note.TargetIds.Add("api");
        envelope.Annotations.Add(note);
        envelope.Annotations.Add(new VisualArtifactInterchangeAnnotation {
            Id = "alt-block",
            Role = VisualArtifactInterchangeAnnotationRole.SequenceBlock,
            Kind = "SequenceBlock:Alt",
            Text = "order accepted",
            StartIndex = 0,
            EndIndex = 1,
            Sequence = new VisualArtifactInterchangeSequenceAnnotation { BlockKind = SequenceArtifactBlockKind.Alt }
        });

        OfficeVisioVisualConversionResult result = envelope.ToOfficeVisio();

        Assert.Equal(OfficeVisioVisualProjectionKind.Sequence, result.Report.Projection);
        Assert.Equal(VisualArtifactKind.Mermaid, result.Report.ArtifactKind);
        Assert.Equal(VisualArtifactInterchangeFamily.Sequence, result.Report.SemanticFamily);
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
        Assert.Equal("SequenceParticipant", api.GetShapeDataValue("CFX.Role"));
        Assert.Equal("Control", api.GetShapeDataValue("CFX.SequenceParticipantKind"));
        Assert.Equal("False", api.GetShapeDataValue("CFX.SequenceParticipantImplicit"));
        Assert.Equal("Commerce", api.GetShapeDataValue("Extension.Owner"));
        Assert.Equal("EU", api.GetShapeDataValue("Detail.1.Region"));
        Assert.Equal("Orders API" + Environment.NewLine + "v2", api.Text);
        Assert.Contains(api.Hyperlinks, link => link.Address == "https://example.test/orders" && link.Description == "Orders runbook");
        Assert.Equal("request", requestConnector.GetShapeDataValue("CFX.Id"));
        Assert.Equal("Async", requestConnector.GetShapeDataValue("CFX.SequenceMessageKind"));
        Assert.Equal("True", requestConnector.GetShapeDataValue("CFX.SequenceActivatesTarget"));
        Assert.Equal("Healthy", requestConnector.GetShapeDataValue("CFX.Status"));
        Assert.Equal("Commerce", requestConnector.GetShapeDataValue("Extension.Owner"));
        Assert.Equal("Create order | async | audited", requestConnector.Label);
        Assert.Equal("client", requestConnector.GetShapeDataValue("CFX.SourceLabel"));
        Assert.Equal("service", requestConnector.GetShapeDataValue("CFX.TargetLabel"));
        Assert.Equal("0", requestConnector.GetShapeDataValue("CFX.Order"));
        Assert.Contains(requestConnector.Hyperlinks, link => link.Address == "https://example.test/create-order" && link.Description == "Create order contract");
        Assert.Equal("Checkout user", result.Page.Shapes.Single(shape => shape.Id == "customer").GetShapeDataValue("CFX.Tooltip"));
        Assert.Equal("Created response", result.Page.Connectors.Single(connector => connector.Id == "response").GetShapeDataValue("CFX.Tooltip"));
        Assert.Contains(result.Report.Warnings, warning => warning.Contains("Sequence participant 'customer'") && warning.Contains("retained as Shape Data"));
        Assert.Contains(result.Report.Warnings, warning => warning.Contains("Sequence message 'response'") && warning.Contains("retained as Shape Data"));
        Assert.True(retryNote.PinX - retryNote.Width / 2D > api.PinX + api.Width / 2D);
    }

    [Fact]
    public void SequenceEqualOrdersPreserveEnvelopeCollectionOrder() {
        VisualArtifactInterchangeEnvelope envelope = SequenceEnvelope("stable-order");
        envelope.Nodes.Add(Participant("z-first", "First", SequenceArtifactParticipantKind.Participant, 0));
        envelope.Nodes.Add(Participant("a-second", "Second", SequenceArtifactParticipantKind.Participant, 0));
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
    public void SequenceProjectionMapsMessageKindsIndependentlyFromLineStyles() {
        VisualArtifactInterchangeEnvelope envelope = SequenceEnvelope("message-directions");
        envelope.Nodes.Add(Participant("caller", "Caller", SequenceArtifactParticipantKind.Actor, 0));
        envelope.Nodes.Add(Participant("service", "Service", SequenceArtifactParticipantKind.Control, 1));
        envelope.Edges.Add(Message("forward", "caller", "service", "Forward", 0));
        envelope.Edges.Add(Message("backward", "caller", "service", "Backward", 1));
        envelope.Edges.Add(Message("both", "caller", "service", "Both", 2));
        envelope.Edges.Add(Message("none", "caller", "service", "None", 3));
        envelope.Edges[1].Sequence!.Kind = SequenceArtifactMessageKind.Return;
        envelope.Edges[2].Sequence!.Kind = SequenceArtifactMessageKind.Async;
        envelope.Edges[2].Sequence!.LineStyle = SequenceArtifactMessageLineStyle.Dashed;
        envelope.Edges[3].Sequence!.Kind = SequenceArtifactMessageKind.Event;

        OfficeVisioVisualConversionResult result = envelope.ToOfficeVisio();

        VisioConnector forward = result.Page.Connectors.Single(connector => connector.Id == "forward");
        VisioConnector backward = result.Page.Connectors.Single(connector => connector.Id == "backward");
        VisioConnector both = result.Page.Connectors.Single(connector => connector.Id == "both");
        VisioConnector none = result.Page.Connectors.Single(connector => connector.Id == "none");
        Assert.Equal(EndArrow.None, forward.BeginArrow);
        Assert.Equal(EndArrow.Triangle, forward.EndArrow);
        Assert.Equal(EndArrow.None, backward.BeginArrow);
        Assert.Equal(EndArrow.Arrow, backward.EndArrow);
        Assert.Equal(EndArrow.None, both.BeginArrow);
        Assert.Equal(EndArrow.Arrow, both.EndArrow);
        Assert.Equal(EndArrow.None, none.BeginArrow);
        Assert.Equal(EndArrow.Arrow, none.EndArrow);
        Assert.Equal(1, backward.LinePattern);
        Assert.Equal(2, both.LinePattern);
    }

    [Fact]
    public void SequencePreservesLateAnnotationRowsAndDeterministicOpenActivations() {
        VisualArtifactInterchangeEnvelope envelope = SequenceEnvelope("late-rows");
        envelope.Nodes.Add(Participant("caller", "Caller", SequenceArtifactParticipantKind.Actor, 0));
        envelope.Nodes.Add(Participant("z-worker", "Z worker", SequenceArtifactParticipantKind.Control, 1));
        envelope.Nodes.Add(Participant("a-worker", "A worker", SequenceArtifactParticipantKind.Control, 2));
        envelope.Edges.Add(Message("activate-z", "caller", "z-worker", "Z", 0, activates: true));
        envelope.Edges.Add(Message("activate-a", "caller", "a-worker", "A", 1, activates: true));
        var note = SequenceNote("late-note", "Later", SequenceArtifactNotePlacement.RightOf, 5);
        note.TargetIds.Add("caller");
        envelope.Annotations.Add(note);
        envelope.Annotations.Add(SequenceBlock("late-block", "Later block", SequenceArtifactBlockKind.Opt, 4, 6));

        OfficeVisioVisualConversionResult result = envelope.ToOfficeVisio();

        Assert.Equal("5", result.Page.Shapes.Single(shape => shape.Id == "late-note").GetUserCellValue("OfficeIMO.SequenceRowIndex"));
        Assert.Equal("4", result.Page.Shapes.Single(shape => shape.Id == "late-block").GetUserCellValue("OfficeIMO.SequenceStartRowIndex"));
        Assert.Equal("6", result.Page.Shapes.Single(shape => shape.Id == "late-block").GetUserCellValue("OfficeIMO.SequenceEndRowIndex"));
        Assert.Equal("a-worker", result.Page.Shapes.Single(shape => shape.Id == "activation-1").GetUserCellValue("OfficeIMO.SequenceParticipantId"));
        Assert.Equal("z-worker", result.Page.Shapes.Single(shape => shape.Id == "activation-2").GetUserCellValue("OfficeIMO.SequenceParticipantId"));
    }

    [Fact]
    public void SequenceProjectsStandaloneTypedActivationEvents() {
        VisualArtifactInterchangeEnvelope envelope = SequenceEnvelope("standalone-activations");
        envelope.Nodes.Add(Participant("caller", "Caller", SequenceArtifactParticipantKind.Actor, 0));
        envelope.Nodes.Add(Participant("service", "Service", SequenceArtifactParticipantKind.Control, 1));
        envelope.Edges.Add(Message("request", "caller", "service", "Request", 0));
        envelope.Edges.Add(Message("response", "service", "caller", "Response", 1, dashed: true));
        envelope.Annotations.Add(SequenceActivation("start", "service", true, 0));
        envelope.Annotations.Add(SequenceActivation("stop", "service", false, 1));

        OfficeVisioVisualConversionResult result = envelope.ToOfficeVisio();

        VisioShape activation = result.Page.Shapes.Single(shape => shape.GetUserCellValue("OfficeIMO.SequenceParticipantId") == "service" &&
            shape.GetUserCellValue("OfficeIMO.SequenceStartRowIndex") == "0" &&
            shape.GetUserCellValue("OfficeIMO.SequenceEndRowIndex") == "1");
        Assert.Equal("service", activation.GetUserCellValue("OfficeIMO.SequenceParticipantId"));
    }

    [Fact]
    public void SequenceMergesInlineAndStandaloneActivationSourcesWithoutDuplicatingEquivalentEvents() {
        VisualArtifactInterchangeEnvelope envelope = SequenceEnvelope("mixed-activations");
        envelope.Nodes.Add(Participant("caller", "Caller", SequenceArtifactParticipantKind.Actor, 0));
        envelope.Nodes.Add(Participant("service", "Service", SequenceArtifactParticipantKind.Control, 1));
        envelope.Nodes.Add(Participant("worker", "Worker", SequenceArtifactParticipantKind.Control, 2));
        envelope.Edges.Add(Message("request", "caller", "service", "Request", 0, activates: true));
        envelope.Edges.Add(Message("dispatch", "caller", "worker", "Dispatch", 1));
        envelope.Edges.Add(Message("response", "service", "caller", "Response", 2, deactivates: true));
        envelope.Edges.Add(Message("complete", "worker", "caller", "Complete", 3));
        envelope.Annotations.Add(SequenceActivation("duplicate-service-start", "service", true, 0));
        envelope.Annotations.Add(SequenceActivation("worker-start", "worker", true, 1));
        envelope.Annotations.Add(SequenceActivation("worker-stop", "worker", false, 3));

        OfficeVisioVisualConversionResult result = envelope.ToOfficeVisio();

        VisioShape[] activations = result.Page.Shapes
            .Where(shape => shape.GetUserCellValue("OfficeIMO.SequenceParticipantId") != null)
            .ToArray();
        Assert.Equal(2, activations.Length);
        Assert.Contains(activations, shape => shape.GetUserCellValue("OfficeIMO.SequenceParticipantId") == "service" &&
            shape.GetUserCellValue("OfficeIMO.SequenceStartRowIndex") == "0" && shape.GetUserCellValue("OfficeIMO.SequenceEndRowIndex") == "2");
        Assert.Contains(activations, shape => shape.GetUserCellValue("OfficeIMO.SequenceParticipantId") == "worker" &&
            shape.GetUserCellValue("OfficeIMO.SequenceStartRowIndex") == "1" && shape.GetUserCellValue("OfficeIMO.SequenceEndRowIndex") == "3");
    }

    [Fact]
    public void SequencePreservesNestedAndLateOpenStandaloneActivations() {
        VisualArtifactInterchangeEnvelope envelope = SequenceEnvelope("nested-late-activations");
        envelope.Nodes.Add(Participant("caller", "Caller", SequenceArtifactParticipantKind.Actor, 0));
        envelope.Nodes.Add(Participant("service", "Service", SequenceArtifactParticipantKind.Control, 1));
        envelope.Edges.Add(Message("request", "caller", "service", "Request", 0));
        envelope.Annotations.Add(SequenceActivation("outer-start", "service", true, 2));
        envelope.Annotations.Add(SequenceActivation("inner-start", "service", true, 3));
        envelope.Annotations.Add(SequenceActivation("inner-stop", "service", false, 4));

        OfficeVisioVisualConversionResult result = envelope.ToOfficeVisio();

        VisioShape[] activations = result.Page.Shapes
            .Where(shape => shape.GetUserCellValue("OfficeIMO.SequenceParticipantId") == "service")
            .ToArray();
        Assert.Equal(2, activations.Length);
        Assert.Contains(activations, shape => shape.GetUserCellValue("OfficeIMO.SequenceStartRowIndex") == "3" &&
            shape.GetUserCellValue("OfficeIMO.SequenceEndRowIndex") == "4");
        Assert.Contains(activations, shape => shape.GetUserCellValue("OfficeIMO.SequenceStartRowIndex") == "2" &&
            shape.GetUserCellValue("OfficeIMO.SequenceEndRowIndex") == "4");
    }

    [Fact]
    public void SequenceTargetlessNotesAreDiagnosedAndSkippedWithoutProjectionFailure() {
        VisualArtifactInterchangeEnvelope envelope = SequenceEnvelope("targetless-note");
        envelope.Nodes.Add(Participant("service", "Service", SequenceArtifactParticipantKind.Control, 0));
        envelope.Annotations.Add(SequenceNote("orphan", "Original participant was removed", SequenceArtifactNotePlacement.RightOf, 0));

        OfficeVisioVisualConversionResult result = envelope.ToOfficeVisio();

        Assert.DoesNotContain(result.Page.Shapes, shape => shape.Id == "orphan");
        Assert.Equal(0, result.Report.AnnotationCount);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == OfficeVisioVisualDiagnosticCode.AnnotationNotProjected &&
            diagnostic.EntityId == "orphan" && diagnostic.Feature == "noteTarget");
    }

    [Fact]
    public void SequenceUnknownNumericParticipantKindFailsValidation() {
        VisualArtifactInterchangeEnvelope envelope = SequenceEnvelope("unknown-kind");
        envelope.Nodes.Add(Participant("participant", "Participant", (SequenceArtifactParticipantKind)999, 0));

        Assert.Throws<ArgumentOutOfRangeException>(() => envelope.ToOfficeVisio());
    }

    [Fact]
    public void ParticipantOnlySequenceProjectsNotesAndFragmentsAtRowZero() {
        VisualArtifactInterchangeEnvelope envelope = SequenceEnvelope("participant-only");
        envelope.Nodes.Add(Participant("service", "Service", SequenceArtifactParticipantKind.Control, 0));
        var note = SequenceNote("note", "Ready", SequenceArtifactNotePlacement.Over, 0);
        note.TargetIds.Add("service");
        envelope.Annotations.Add(note);
        var fragment = SequenceBlock("fragment", "cached", SequenceArtifactBlockKind.Opt, 0, 0);
        fragment.TargetIds.Add("service");
        envelope.Annotations.Add(fragment);

        OfficeVisioVisualConversionResult result = envelope.ToOfficeVisio();

        Assert.Equal(2, result.Report.AnnotationCount);
        Assert.Contains(result.Page.Shapes, shape => shape.Id == "note");
        Assert.Contains(result.Page.Shapes, shape => shape.Id == "fragment");
        Assert.Contains(result.Report.Warnings, warning => warning.Contains("note") && warning.Contains("Over") && warning.Contains("right-side note placement"));
    }

    [Fact]
    public void SequencePageUsesEnvelopeDimensionsOnlyWhenNaturalSizingIsRequested() {
        VisualArtifactInterchangeEnvelope envelope = SequenceEnvelope("wide-sequence", width: 2400, height: 1200);
        envelope.Nodes.Add(Participant("caller", "Caller", SequenceArtifactParticipantKind.Actor, 0));
        envelope.Nodes.Add(Participant("service", "Service", SequenceArtifactParticipantKind.Control, 1));
        envelope.Edges.Add(Message("call", "caller", "service", "Call", 0));
        var note = SequenceNote("note", "Natural page note", SequenceArtifactNotePlacement.RightOf, 0);
        note.TargetIds.Add("service");
        envelope.Annotations.Add(note);

        OfficeVisioVisualConversionResult fitted = envelope.ToOfficeVisio();
        OfficeVisioVisualConversionResult natural = envelope.ToOfficeVisio(new OfficeVisioVisualOptions { UseNaturalPageSize = true, PixelsPerInch = 100D });

        Assert.True(fitted.Page.Width < natural.Page.Width);
        Assert.True(fitted.Page.Width < 9D);
        Assert.True(natural.Page.Width >= 24D);
        VisioShapeBounds naturalBounds = natural.Page.GetContentBounds();
        Assert.True(Math.Abs(naturalBounds.CenterX - natural.Page.Width / 2D) < 0.001D);
        Assert.True(Math.Abs(naturalBounds.CenterY - natural.Page.Height / 2D) < 0.001D);
    }

    [Fact]
    public void CompactSequenceNaturalViewportGrowsOnlyWhenContentRequiresIt() {
        VisualArtifactInterchangeEnvelope envelope = SequenceEnvelope("compact-sequence", width: 400, height: 300);
        envelope.Nodes.Add(Participant("service", "Service", SequenceArtifactParticipantKind.Control, 0));

        OfficeVisioVisualConversionResult result = envelope.ToOfficeVisio(new OfficeVisioVisualOptions {
            UseNaturalPageSize = true,
            PixelsPerInch = 100D
        });

        Assert.InRange(result.Page.Width, 4D, 4.5D);
        Assert.InRange(result.Page.Height, 3D, 4D);
        VisioShapeBounds bounds = result.Page.GetContentBounds();
        Assert.True(Math.Abs(bounds.CenterX - result.Page.Width / 2D) < 0.001D);
        Assert.True(Math.Abs(bounds.CenterY - result.Page.Height / 2D) < 0.001D);
    }

    [Fact]
    public void SourceGraphIdCollisionsAreRemappedByTheInterchangeOwnerBeforeVisioProjection() {
        TopologyChart topology = TopologyChart.Create();
        topology.LayoutMode = TopologyLayoutMode.Manual;
        topology.Groups.Add(new TopologyGroup { Id = "shared", Label = "Group", X = 0, Y = 0, Width = 400, Height = 200 });
        topology.Nodes.Add(new TopologyNode { Id = "shared", Label = "Source", GroupId = "shared", X = 30, Y = 70, Width = 100, Height = 50 });
        topology.Nodes.Add(new TopologyNode { Id = "target", Label = "Target", GroupId = "shared", X = 220, Y = 70, Width = 100, Height = 50 });
        topology.Edges.Add(new TopologyEdge { Id = "shared", SourceNodeId = "shared", TargetNodeId = "target" });

        OfficeVisioVisualConversionResult result = topology.ToVisualArtifact().ToOfficeVisio();

        Assert.Equal(result.Envelope.Groups.Count + result.Envelope.Nodes.Count + result.Envelope.Edges.Count,
            result.Envelope.Groups.Select(group => group.Id)
                .Concat(result.Envelope.Nodes.Select(node => node.Id))
                .Concat(result.Envelope.Edges.Select(edge => edge.Id))
                .Distinct(StringComparer.Ordinal)
                .Count());
        Assert.Equal(result.Envelope.Nodes.Single(node => node.Label == "Source").Id,
            result.Envelope.Edges.Single().SourceId);
    }

    [Fact]
    public void SequenceSemanticIdsAreMappedAwayFromGeneratedVisioHelpers() {
        VisualArtifactInterchangeEnvelope envelope = SequenceEnvelope("collisions", "Collisions");
        envelope.Nodes.Add(Participant("api", "API", SequenceArtifactParticipantKind.Control, 0));
        envelope.Nodes.Add(Participant("api-lifeline", "Worker", SequenceArtifactParticipantKind.Participant, 1));
        envelope.Edges.Add(Message("api-lifeline-end", "api", "api-lifeline", "Dispatch", 0, activates: true));
        var fragment = new VisualArtifactInterchangeAnnotation {
            Id = "message-api-lifeline-end-from",
            Role = VisualArtifactInterchangeAnnotationRole.SequenceBlock,
            Kind = "SequenceBlock:Opt",
            Text = "mapped",
            StartIndex = 0,
            EndIndex = 0,
            Sequence = new VisualArtifactInterchangeSequenceAnnotation { BlockKind = SequenceArtifactBlockKind.Opt }
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
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == OfficeVisioVisualDiagnosticCode.IdRemapped &&
            diagnostic.Severity == OfficeVisioVisualDiagnosticSeverity.Information);
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
    public void TypedSequenceBlockRoleReservesNativeFragmentHelpersIndependentlyFromFreeFormKind() {
        VisualArtifactInterchangeEnvelope envelope = SequenceEnvelope("typed-fragment-id");
        envelope.Nodes.Add(Participant("fragment-label", "Reserved helper", SequenceArtifactParticipantKind.Actor, 0));
        envelope.Nodes.Add(Participant("service", "Service", SequenceArtifactParticipantKind.Control, 1));
        envelope.Edges.Add(Message("call", "fragment-label", "service", "Call", 0));
        VisualArtifactInterchangeAnnotation fragment = SequenceBlock("fragment", "Optional", SequenceArtifactBlockKind.Opt, 0, 0);
        fragment.Kind = "custom-block-kind";
        fragment.TargetIds.Add("service");
        envelope.Annotations.Add(fragment);

        OfficeVisioVisualConversionResult result = envelope.ToOfficeVisio();

        Assert.Contains(result.Page.Shapes, shape => shape.GetShapeDataValue("CFX.Id") == "fragment");
        Assert.Equal(result.Page.Shapes.Count, result.Page.Shapes.Select(shape => shape.Id).Distinct(StringComparer.Ordinal).Count());
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == OfficeVisioVisualDiagnosticCode.IdRemapped &&
            diagnostic.EntityKind == OfficeVisioVisualEntityKind.Annotation && diagnostic.EntityId == "fragment");
    }

    [Fact]
    public void TypedSequenceBranchesReserveAllNativeOperandHelpers() {
        VisualArtifactInterchangeEnvelope envelope = SequenceEnvelope("typed-branch-ids");
        envelope.Nodes.Add(Participant("branch-label", "Guard helper", SequenceArtifactParticipantKind.Actor, 0));
        envelope.Nodes.Add(Participant("partition-from", "Divider start", SequenceArtifactParticipantKind.Participant, 1));
        envelope.Nodes.Add(Participant("partition-to", "Divider end", SequenceArtifactParticipantKind.Control, 2));
        envelope.Edges.Add(Message("first", "branch-label", "partition-from", "First", 0));
        envelope.Edges.Add(Message("second", "partition-from", "partition-to", "Second", 1));
        envelope.Annotations.Add(SequenceBlock("fragment", "Choice", SequenceArtifactBlockKind.Alt, 0, 1));
        envelope.Annotations.Add(SequenceBranch("branch", "Primary", SequenceArtifactBlockKind.Alt, "Primary", 0, 0, 0));
        envelope.Annotations.Add(SequenceBranch("partition", "Alternate", SequenceArtifactBlockKind.Alt, "Else", 1, 1, 0));

        OfficeVisioVisualConversionResult result = envelope.ToOfficeVisio();

        Assert.Equal(result.Page.Shapes.Count, result.Page.Shapes.Select(shape => shape.Id).Distinct(StringComparer.Ordinal).Count());
        Assert.Equal(result.Page.Connectors.Count, result.Page.Connectors.Select(connector => connector.Id).Distinct(StringComparer.Ordinal).Count());
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == OfficeVisioVisualDiagnosticCode.IdRemapped &&
            diagnostic.EntityKind == OfficeVisioVisualEntityKind.Annotation && diagnostic.EntityId == "branch");
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == OfficeVisioVisualDiagnosticCode.IdRemapped &&
            diagnostic.EntityKind == OfficeVisioVisualEntityKind.Annotation && diagnostic.EntityId == "partition");
    }

    [Fact]
    public void GraphProjectionPreservesForwardBackwardAndBidirectionalArrows() {
        VisualArtifactInterchangeEnvelope envelope = CreateTopologyEnvelope();
        envelope.Edges.Clear();
        envelope.Edges.Add(TopologyEdge("forward", VisualLinkDirection.Forward));
        envelope.Edges.Add(TopologyEdge("backward", VisualLinkDirection.Backward));
        envelope.Edges.Add(TopologyEdge("both", VisualLinkDirection.Bidirectional));

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
    public void GraphProjectionPreservesExplicitLineStylesAndStandaloneTooltips() {
        VisualArtifactInterchangeEnvelope envelope = CreateTopologyEnvelope();
        envelope.Nodes.Single(node => node.Id == "api").Href = null;
        envelope.Nodes.Single(node => node.Id == "api").Color = "#112233";
        envelope.Nodes.Single(node => node.Id == "api").BackgroundColor = "rgb(68 85 102)";
        envelope.Nodes.Single(node => node.Id == "database").Tooltip = "Database owner";
        envelope.Groups.Single().Tooltip = "Data boundary";
        envelope.Groups.Single().Color = "#778899";
        envelope.Edges.Clear();
        envelope.Edges.Add(TopologyEdge("solid", VisualLinkDirection.Forward, TopologyEdgeLineStyle.Solid, "#AABBCC"));
        VisualArtifactInterchangeEdge dashed = TopologyEdge("dashed", VisualLinkDirection.Forward, TopologyEdgeLineStyle.Dashed);
        dashed.Tooltip = "Retry path";
        dashed.SourceLabel = "client";
        dashed.TargetLabel = "service";
        envelope.Edges.Add(dashed);
        envelope.Edges.Add(TopologyEdge("dotted", VisualLinkDirection.Forward, TopologyEdgeLineStyle.Dotted));

        OfficeVisioVisualConversionResult result = envelope.ToOfficeVisio();

        Assert.Equal(1, result.Page.Connectors.Single(connector => connector.Id == "solid").LinePattern);
        Assert.Equal(2, result.Page.Connectors.Single(connector => connector.Id == "dashed").LinePattern);
        Assert.Equal(3, result.Page.Connectors.Single(connector => connector.Id == "dotted").LinePattern);
        Assert.Equal(Color.Parse("#112233"), result.Page.Shapes.Single(shape => shape.Id == "api").LineColor);
        Assert.Equal(Color.Parse("#445566"), result.Page.Shapes.Single(shape => shape.Id == "api").FillColor);
        Assert.Equal(Color.Parse("#778899"), result.Page.Shapes.Single(shape => shape.Id == "data-zone").LineColor);
        Assert.Equal(Color.Parse("#AABBCC"), result.Page.Connectors.Single(connector => connector.Id == "solid").LineColor);
        Assert.Equal("API runbook", result.Page.Shapes.Single(shape => shape.Id == "api").GetShapeDataValue("CFX.Tooltip"));
        Assert.Equal("Database owner", result.Page.Shapes.Single(shape => shape.Id == "database").GetShapeDataValue("CFX.Tooltip"));
        Assert.Equal("Data boundary", result.Page.Shapes.Single(shape => shape.Id == "data-zone").GetShapeDataValue("CFX.Tooltip"));
        Assert.Equal("Retry path", result.Page.Connectors.Single(connector => connector.Id == "dashed").GetShapeDataValue("CFX.Tooltip"));
        Assert.Contains(result.Report.Warnings, warning => warning.Contains("Node 'api'") && warning.Contains("retained as Shape Data"));
        Assert.Contains(result.Report.Warnings, warning => warning.Contains("Group 'data-zone'") && warning.Contains("retained as Shape Data"));
        Assert.Contains(result.Report.Warnings, warning => warning.Contains("Edge 'dashed'") && warning.Contains("retained as Shape Data"));
        Assert.Contains(result.Report.Warnings, warning => warning.Contains("Edge 'dashed'") && warning.Contains("endpoint labels") && warning.Contains("retained as Shape Data"));

        OfficeVisioVisualConversionResult withoutShapeData = envelope.ToOfficeVisio(new OfficeVisioVisualOptions { IncludeShapeData = false });
        Assert.Contains(withoutShapeData.Report.Warnings, warning => warning.Contains("Node 'api'") && warning.Contains("remains only in the CFX envelope"));
        Assert.Contains(withoutShapeData.Report.Warnings, warning => warning.Contains("Group 'data-zone'") && warning.Contains("remains only in the CFX envelope"));
        Assert.Contains(withoutShapeData.Report.Warnings, warning => warning.Contains("Edge 'dashed'") && warning.Contains("remains only in the CFX envelope"));
        Assert.Contains(withoutShapeData.Report.Warnings, warning => warning.Contains("Edge 'dashed'") && warning.Contains("endpoint labels") && warning.Contains("remain only in the CFX envelope"));

        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
        try {
            result.Document.Save(path);
            Assert.Empty(VisioValidator.Validate(path));
            VisioDocument loaded = VisioDocument.Load(path);
            Assert.Equal(3, loaded.Pages[0].Connectors.Single(connector => connector.Id == "dotted").LinePattern);
            Assert.Equal("Database owner", loaded.Pages[0].Shapes.Single(shape => shape.Id == "database").GetShapeDataValue("CFX.Tooltip"));
            Assert.Equal("Retry path", loaded.Pages[0].Connectors.Single(connector => connector.Id == "dashed").GetShapeDataValue("CFX.Tooltip"));
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void SequenceProjectionPreservesTypedSolidAndDashedMessageStyles() {
        VisualArtifactInterchangeEnvelope envelope = SequenceEnvelope("styled-sequence");
        VisualArtifactInterchangeNode caller = Participant("caller", "Caller", SequenceArtifactParticipantKind.Actor, 0);
        caller.Color = "#123456";
        caller.BackgroundColor = "#ABCDEF";
        envelope.Nodes.Add(caller);
        envelope.Nodes.Add(Participant("service", "Service", SequenceArtifactParticipantKind.Control, 1));
        envelope.Edges.Add(Message("solid", "caller", "service", "Solid", 0));
        VisualArtifactInterchangeEdge dashedMessage = Message("dashed", "service", "caller", "Dashed", 1, dashed: true);
        dashedMessage.Color = "#336699";
        envelope.Edges.Add(dashedMessage);

        OfficeVisioVisualConversionResult result = envelope.ToOfficeVisio();

        Assert.Equal(1, result.Page.Connectors.Single(connector => connector.Id == "solid").LinePattern);
        Assert.Equal(2, result.Page.Connectors.Single(connector => connector.Id == "dashed").LinePattern);
        Assert.Equal(Color.Parse("#336699"), result.Page.Connectors.Single(connector => connector.Id == "dashed").LineColor);
        Assert.Equal(Color.Parse("#123456"), result.Page.Shapes.Single(shape => shape.Id == "caller").LineColor);
        Assert.Equal(Color.Parse("#ABCDEF"), result.Page.Shapes.Single(shape => shape.Id == "caller").FillColor);

        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
        try {
            result.Document.Save(path);
            Assert.Empty(VisioValidator.Validate(path));
            VisioDocument loaded = VisioDocument.Load(path);
            Assert.Equal(2, loaded.Pages[0].Connectors.Single(connector => connector.Id == "dashed").LinePattern);
            Assert.Equal(Color.Parse("#336699"), loaded.Pages[0].Connectors.Single(connector => connector.Id == "dashed").LineColor);
            Assert.Equal(Color.Parse("#123456"), loaded.Pages[0].Shapes.Single(shape => shape.Id == "caller").LineColor);
            Assert.Equal(Color.Parse("#ABCDEF"), loaded.Pages[0].Shapes.Single(shape => shape.Id == "caller").FillColor);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void TopologyLayoutTokensMapToNativeStrategiesAndReportUnsupportedLayouts() {
        var matrix = TopologyEnvelope("matrix", TopologyLayoutMode.Matrix);
        for (int index = 0; index < 4; index++) matrix.Nodes.Add(TopologyNode("node-" + index, "Node " + index));

        OfficeVisioVisualConversionResult matrixResult = matrix.ToOfficeVisio();
        Assert.True(matrixResult.Page.Shapes.Where(shape => shape.Id.StartsWith("node-", StringComparison.Ordinal)).Select(shape => shape.PinX).Distinct().Count() > 1);
        Assert.True(matrixResult.Page.Shapes.Where(shape => shape.Id.StartsWith("node-", StringComparison.Ordinal)).Select(shape => shape.PinY).Distinct().Count() > 1);

        matrix.Topology!.LayoutMode = TopologyLayoutMode.Geographic;
        matrix.Topology.LayoutDirection = TopologyLayoutDirection.RightToLeft;
        OfficeVisioVisualConversionResult geographicResult = matrix.ToOfficeVisio();
        Assert.Contains(geographicResult.Report.Diagnostics, diagnostic => diagnostic.Code == OfficeVisioVisualDiagnosticCode.LayoutNormalized);
        Assert.Contains(geographicResult.Report.Diagnostics, diagnostic => diagnostic.Code == OfficeVisioVisualDiagnosticCode.DirectionNormalized);
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
    public void SequenceMetadataLossIsReportedWhenShapeDataProjectionIsDisabled() {
        VisualArtifactInterchangeEnvelope envelope = SequenceEnvelope("metadata-sequence");
        envelope.Extensions["Owner"] = "Platform";
        VisualArtifactInterchangeNode service = Participant("service", "Service", SequenceArtifactParticipantKind.Control, 0);
        service.X = 100D;
        service.Width = 160D;
        service.Ports.Add(new VisualArtifactInterchangePort { Id = "request", Side = TopologyEdgePort.Left });
        envelope.Nodes.Add(service);

        OfficeVisioVisualConversionResult result = envelope.ToOfficeVisio(new OfficeVisioVisualOptions { IncludeShapeData = false });

        Assert.Contains(result.Report.Warnings, warning => warning.Contains("Sequence-level extensions") && warning.Contains("Shape Data projection was disabled"));
        Assert.Contains(result.Report.Warnings, warning => warning.Contains("participant coordinates and dimensions") && warning.Contains("recomputed"));
        Assert.Contains(result.Report.Warnings, warning => warning.Contains("participant ports") && warning.Contains("remain only in the CFX envelope") && warning.Contains("Shape Data projection was disabled"));
        Assert.True(result.Report.HasSemanticLoss);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == OfficeVisioVisualDiagnosticCode.ShapeDataDisabled);
    }

    [Fact]
    public void DisabledShapeDataAndHyperlinksAreReportedAsSemanticLoss() {
        VisualArtifactInterchangeEnvelope envelope = CreateTopologyEnvelope();
        VisualArtifactInterchangeNode api = envelope.Nodes.Single(node => node.Id == "api");
        api.Href = "https://example.test/api";
        api.Extensions["Owner"] = "Platform";
        api.Metrics.Add(new VisualArtifactInterchangeMetric { Name = "Latency", Value = "42 ms" });

        OfficeVisioVisualConversionResult result = envelope.ToOfficeVisio(new OfficeVisioVisualOptions {
            IncludeShapeData = false,
            IncludeHyperlinks = false
        });

        Assert.True(result.Report.HasSemanticLoss);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == OfficeVisioVisualDiagnosticCode.ShapeDataDisabled &&
            diagnostic.Severity == OfficeVisioVisualDiagnosticSeverity.Warning);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == OfficeVisioVisualDiagnosticCode.HyperlinkNotProjected &&
            diagnostic.EntityId == "api" && diagnostic.Severity == OfficeVisioVisualDiagnosticSeverity.Warning);
    }

    [Fact]
    public void TypedVisioProjectionReportsRenderWatermarksThatRemainInStaticFallbacks() {
        TopologyChart topology = TopologyChart.Create();
        topology.Nodes.Add(new TopologyNode { Id = "service", Label = "Service" });
        var renderOptions = new VisualArtifactRenderOptions();
        renderOptions.Watermarks.Add(VisualWatermark.FromText("CONFIDENTIAL"));

        OfficeVisioVisualConversionResult result = topology.ToVisualArtifact().ToOfficeVisio(renderOptions: renderOptions);

        Assert.Contains(result.Report.Warnings, warning => warning.Contains("watermarks") && warning.Contains("SVG or PNG"));
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

    [Theory]
    [InlineData(TopologyLayoutDirection.LeftToRight)]
    [InlineData(TopologyLayoutDirection.TopToBottom)]
    public void CompactGraphNaturalViewportGrowsOnlyToContentAndCentersIt(TopologyLayoutDirection direction) {
        VisualArtifactInterchangeEnvelope envelope = TopologyEnvelope("compact-graph", TopologyLayoutMode.Layered, direction);
        envelope.Width = 192;
        envelope.Height = 144;
        envelope.Nodes.Add(TopologyNode("service", "Service"));

        OfficeVisioVisualConversionResult result = envelope.ToOfficeVisio(new OfficeVisioVisualOptions {
            UseNaturalPageSize = true,
            PixelsPerInch = 96D
        });

        Assert.InRange(result.Page.Width, 3.2D, 3.5D);
        Assert.InRange(result.Page.Height, 2.4D, 2.7D);
        VisioShapeBounds bounds = result.Page.GetContentBounds();
        Assert.True(Math.Abs(bounds.CenterX - result.Page.Width / 2D) < 0.001D);
        Assert.True(Math.Abs(bounds.CenterY - result.Page.Height / 2D) < 0.001D);
    }

    [Fact]
    public void FidelityCountsOnlyNativeObjectsAndWarnsForUnmappedGraphAnnotations() {
        VisualArtifactInterchangeEnvelope envelope = CreateTopologyEnvelope();
        envelope.Extensions["Owner"] = "Platform";
        envelope.Annotations.Add(new VisualArtifactInterchangeAnnotation {
            Id = "graph-note",
            Kind = "Note",
            Text = "Retained only in the semantic envelope"
        });

        OfficeVisioVisualConversionResult result = envelope.ToOfficeVisio(new OfficeVisioVisualOptions { IncludeGroups = false });

        Assert.Equal(0, result.Report.GroupCount);
        Assert.Equal(0, result.Report.AnnotationCount);
        Assert.Contains(result.Report.Warnings, warning => warning.Contains("graph-note"));
        Assert.Contains(result.Report.Warnings, warning => warning.Contains("Artifact-level extensions"));
        Assert.Contains(result.Report.Warnings, warning => warning.Contains("native Visio graph layout selected connector sides"));
        VisioConnector edge = result.Page.Connectors.Single(connector => connector.Id == "api-db");
        Assert.Equal("out", edge.GetShapeDataValue("CFX.SourcePortId"));
        Assert.Equal("0", edge.GetShapeDataValue("CFX.Order"));
    }

    [Fact]
    public void SequenceFragmentsPreserveNestingBranchesAndTypedShapeData() {
        VisualArtifactInterchangeEnvelope envelope = SequenceEnvelope("nested-fragments");
        envelope.Nodes.Add(Participant("client", "Client", SequenceArtifactParticipantKind.Actor, 0));
        envelope.Nodes.Add(Participant("service", "Service", SequenceArtifactParticipantKind.Control, 1));
        envelope.Edges.Add(Message("request", "client", "service", "Request", 0));
        envelope.Edges.Add(Message("retry", "service", "service", "Retry", 1));
        envelope.Edges.Add(Message("response", "service", "client", "Response", 2));
        envelope.Annotations.Add(SequenceBlock("outer", "Accepted", SequenceArtifactBlockKind.Alt, 0, 2));
        envelope.Annotations.Add(SequenceBlock("inner", "Retry enabled", SequenceArtifactBlockKind.Opt, 1, 1));
        envelope.Annotations.Add(SequenceBranch("primary", "Accepted", SequenceArtifactBlockKind.Alt, "Primary", 0, 0, 0));
        envelope.Annotations.Add(SequenceBranch("fallback", "Rejected", SequenceArtifactBlockKind.Alt, "Else", 1, 2, 0));

        OfficeVisioVisualConversionResult result = envelope.ToOfficeVisio();

        VisioShape outer = Assert.Single(result.Page.Shapes, shape => shape.Id == "outer");
        VisioShape inner = Assert.Single(result.Page.Shapes, shape => shape.Id == "inner");
        Assert.Equal("outer", inner.GetUserCellValue("OfficeIMO.SequenceParentFragmentId"));
        VisioShape primary = Assert.Single(result.Page.Shapes, shape => shape.GetUserCellValue("OfficeIMO.SequenceFragmentOperandId") == "primary");
        VisioShape fallback = Assert.Single(result.Page.Shapes, shape => shape.GetUserCellValue("OfficeIMO.SequenceFragmentOperandId") == "fallback");
        Assert.Equal("Primary", primary.GetShapeDataValue("CFX.SequenceBranchKind"));
        Assert.Equal("Else", fallback.GetShapeDataValue("CFX.SequenceBranchKind"));
        Assert.Equal("Alt", outer.GetShapeDataValue("CFX.SequenceBlockKind"));
        Assert.False(result.Report.HasSemanticLoss);
    }

    [Fact]
    public void StandaloneActivationEventsRemainAssociatedWithTheirNativeActivation() {
        VisualArtifactInterchangeEnvelope envelope = SequenceEnvelope("activation-metadata");
        envelope.Nodes.Add(Participant("client", "Client", SequenceArtifactParticipantKind.Actor, 0));
        envelope.Nodes.Add(Participant("service", "Service", SequenceArtifactParticipantKind.Control, 1));
        envelope.Edges.Add(Message("request", "client", "service", "Request", 0, activates: true));
        envelope.Edges.Add(Message("response", "service", "client", "Response", 1));
        VisualArtifactInterchangeAnnotation start = SequenceActivation("activation-start", "service", true, 0);
        start.Extensions["Source"] = "Mermaid";
        VisualArtifactInterchangeAnnotation stop = SequenceActivation("activation-stop", "service", false, 1);
        stop.Extensions["Reason"] = "Complete";
        envelope.Annotations.Add(start);
        envelope.Annotations.Add(stop);

        OfficeVisioVisualConversionResult result = envelope.ToOfficeVisio();

        VisioShape activation = Assert.Single(result.Page.Shapes, shape => shape.Id == "activation-start");
        Assert.Equal("activation-start,activation-stop", activation.GetShapeDataValue("CFX.ActivationEventIds"));
        Assert.Equal("Mermaid", activation.GetShapeDataValue("CFX.ActivationEvent.1.Extension.Source"));
        Assert.Equal("Complete", activation.GetShapeDataValue("CFX.ActivationEvent.2.Extension.Reason"));
        Assert.Equal("False", activation.GetShapeDataValue("CFX.ActivationEvent.2.State"));
    }

    [Fact]
    public void GraphSelfEdgesUseTheReusableExternalLoopRoute() {
        VisualArtifactInterchangeEnvelope envelope = TopologyEnvelope("self-edge");
        envelope.Nodes.Add(TopologyNode("api", "API"));
        VisualArtifactInterchangeEdge edge = TopologyEdge("retry", VisualLinkDirection.Forward);
        edge.SourceId = "api";
        edge.TargetId = "api";
        envelope.Edges.Add(edge);

        OfficeVisioVisualConversionResult result = envelope.ToOfficeVisio();

        VisioShape api = Assert.Single(result.Page.Shapes, shape => shape.Id == "api");
        VisioConnector retry = Assert.Single(result.Page.Connectors, connector => connector.Id == "retry");
        Assert.Contains(retry.Waypoints, point => point.X > api.PinX + (api.Width / 2D));
        Assert.Contains(retry.Waypoints, point => point.Y > api.PinY + (api.Height / 2D));
    }

    [Fact]
    public void AlphaBearingGraphAndSequenceColorsAreReportedInsteadOfMadeOpaque() {
        VisualArtifactInterchangeEnvelope graph = TopologyEnvelope("alpha-graph");
        var group = new VisualArtifactInterchangeGroup {
            Id = "zone", Role = VisualArtifactInterchangeGroupRole.TopologyGroup, Kind = "TopologyGroup", Label = "Zone",
            Color = "#11223380", Topology = new VisualArtifactInterchangeTopologyGroup()
        };
        graph.Groups.Add(group);
        VisualArtifactInterchangeNode api = TopologyNode("api", "API");
        api.GroupId = "zone";
        api.Color = "rgba(10, 20, 30, 0.5)";
        api.BackgroundColor = "transparent";
        graph.Nodes.Add(api);
        VisualArtifactInterchangeEdge retry = TopologyEdge("retry", VisualLinkDirection.Forward, color: "#44556680");
        retry.SourceId = "api";
        retry.TargetId = "api";
        graph.Edges.Add(retry);

        OfficeVisioVisualConversionResult graphResult = graph.ToOfficeVisio();
        Assert.Contains(graphResult.Report.Diagnostics, diagnostic => diagnostic.Code == OfficeVisioVisualDiagnosticCode.ColorNotProjected && diagnostic.EntityId == "zone" && diagnostic.Feature == "colorAlpha");
        Assert.Contains(graphResult.Report.Diagnostics, diagnostic => diagnostic.Code == OfficeVisioVisualDiagnosticCode.ColorNotProjected && diagnostic.EntityId == "api" && diagnostic.Feature == "colorAlpha");
        Assert.Contains(graphResult.Report.Diagnostics, diagnostic => diagnostic.Code == OfficeVisioVisualDiagnosticCode.ColorNotProjected && diagnostic.EntityId == "retry" && diagnostic.Feature == "colorAlpha");

        VisualArtifactInterchangeEnvelope sequence = SequenceEnvelope("alpha-sequence");
        VisualArtifactInterchangeNode client = Participant("client", "Client", SequenceArtifactParticipantKind.Actor, 0);
        client.Color = "#11223380";
        sequence.Nodes.Add(client);
        sequence.Nodes.Add(Participant("service", "Service", SequenceArtifactParticipantKind.Control, 1));
        VisualArtifactInterchangeEdge message = Message("call", "client", "service", "Call", 0);
        message.Color = "transparent";
        sequence.Edges.Add(message);
        OfficeVisioVisualConversionResult sequenceResult = sequence.ToOfficeVisio();
        Assert.Contains(sequenceResult.Report.Diagnostics, diagnostic => diagnostic.Code == OfficeVisioVisualDiagnosticCode.ColorNotProjected && diagnostic.EntityId == "client" && diagnostic.Feature == "colorAlpha");
        Assert.Contains(sequenceResult.Report.Diagnostics, diagnostic => diagnostic.Code == OfficeVisioVisualDiagnosticCode.ColorNotProjected && diagnostic.EntityId == "call" && diagnostic.Feature == "colorAlpha");
    }

    private static VisualArtifactInterchangeEnvelope CreateTopologyEnvelope() {
        VisualArtifactInterchangeEnvelope envelope = TopologyEnvelope("service-topology");
        envelope.Title = "Service topology";
        envelope.Width = 900;
        envelope.Height = 520;
        envelope.Groups.Add(new VisualArtifactInterchangeGroup {
            Id = "data-zone",
            Role = VisualArtifactInterchangeGroupRole.TopologyGroup,
            Kind = "TopologyGroup",
            Label = "Data zone",
            Topology = new VisualArtifactInterchangeTopologyGroup()
        });
        VisualArtifactInterchangeNode api = TopologyNode("api", "API", TopologyNodeKind.Service, TopologyHealthStatus.Healthy);
        api.Href = "https://example.test/api";
        api.Tooltip = "API runbook";
        api.Extensions["Owner"] = "Platform";
        api.Extensions["owner"] = "Secondary";
        api.Extensions["Metric.latency"] = "caller-owned";
        api.Metrics.Add(new VisualArtifactInterchangeMetric { Name = "latency", Value = "42 ms" });
        var detail = new VisualArtifactInterchangeDetail { Label = "Port", Value = "443", IconId = "health", Status = "Healthy", Color = "#22AA66" };
        detail.Extensions["Protocol"] = "TCP";
        api.Details.Add(detail);
        var port = new VisualArtifactInterchangePort { Id = "out", Side = TopologyEdgePort.Right, Offset = 0.5D, Label = "egress" };
        port.Extensions["Role"] = "primary";
        api.Ports.Add(port);
        envelope.Nodes.Add(api);
        VisualArtifactInterchangeNode database = TopologyNode("database", "Database", TopologyNodeKind.Database);
        database.GroupId = "data-zone";
        envelope.Nodes.Add(database);
        VisualArtifactInterchangeEdge edge = TopologyEdge("api-db", VisualLinkDirection.Forward);
        edge.Label = "queries";
        edge.SourcePortId = "out";
        edge.Topology!.Kind = TopologyEdgeKind.DataFlow;
        envelope.Edges.Add(edge);
        return envelope;
    }

    private static VisualArtifactInterchangeEnvelope TopologyEnvelope(
        string id,
        TopologyLayoutMode layoutMode = TopologyLayoutMode.Layered,
        TopologyLayoutDirection direction = TopologyLayoutDirection.LeftToRight) => new() {
            Id = id,
            Kind = VisualArtifactKind.Topology,
            Family = VisualArtifactInterchangeFamily.Topology,
            Topology = new VisualArtifactInterchangeTopologyArtifact { LayoutMode = layoutMode, LayoutDirection = direction }
        };

    private static VisualArtifactInterchangeEnvelope SequenceEnvelope(
        string id,
        string? title = null,
        double? width = null,
        double? height = null) => new() {
            Id = id,
            Kind = VisualArtifactKind.Sequence,
            Family = VisualArtifactInterchangeFamily.Sequence,
            Sequence = new VisualArtifactInterchangeSequenceArtifact(),
            Title = title ?? string.Empty,
            Width = width,
            Height = height
        };

    private static VisualArtifactInterchangeNode TopologyNode(
        string id,
        string label,
        TopologyNodeKind kind = TopologyNodeKind.Generic,
        TopologyHealthStatus status = TopologyHealthStatus.Unknown) => new() {
            Id = id,
            Role = VisualArtifactInterchangeNodeRole.TopologyNode,
            Kind = kind.ToString(),
            Label = label,
            Status = status.ToString(),
            Topology = new VisualArtifactInterchangeTopologyNode {
                Kind = kind,
                Status = status,
                DisplayMode = TopologyNodeDisplayMode.Card
            }
        };

    private static VisualArtifactInterchangeEdge TopologyEdge(
        string id,
        VisualLinkDirection direction,
        TopologyEdgeLineStyle lineStyle = TopologyEdgeLineStyle.Solid,
        string? color = null) => new() {
            Id = id,
            Role = VisualArtifactInterchangeEdgeRole.TopologyEdge,
            Kind = TopologyEdgeKind.Generic.ToString(),
            SourceId = "api",
            TargetId = "database",
            Color = color,
            Topology = new VisualArtifactInterchangeTopologyEdge {
                Kind = TopologyEdgeKind.Generic,
                Status = TopologyHealthStatus.Unknown,
                Direction = direction,
                LineStyle = lineStyle,
                Routing = TopologyEdgeRouting.Orthogonal
            }
        };

    private static VisualArtifactInterchangeNode Participant(string id, string label, SequenceArtifactParticipantKind kind, int order) => new() {
        Id = id,
        Role = VisualArtifactInterchangeNodeRole.SequenceParticipant,
        Label = label,
        Kind = kind.ToString(),
        Sequence = new VisualArtifactInterchangeSequenceNode { Kind = kind, Order = order }
    };

    private static VisualArtifactInterchangeAnnotation SequenceNote(
        string id,
        string text,
        SequenceArtifactNotePlacement placement,
        int row) => new() {
            Id = id,
            Role = VisualArtifactInterchangeAnnotationRole.SequenceNote,
            Kind = "SequenceNote",
            Text = text,
            Placement = placement.ToString(),
            StartIndex = row,
            EndIndex = row,
            Sequence = new VisualArtifactInterchangeSequenceAnnotation { NotePlacement = placement }
        };

    private static VisualArtifactInterchangeAnnotation SequenceBlock(
        string id,
        string text,
        SequenceArtifactBlockKind kind,
        int start,
        int end) => new() {
            Id = id,
            Role = VisualArtifactInterchangeAnnotationRole.SequenceBlock,
            Kind = "SequenceBlock:" + kind,
            Text = text,
            StartIndex = start,
            EndIndex = end,
            Sequence = new VisualArtifactInterchangeSequenceAnnotation { BlockKind = kind }
        };

    private static VisualArtifactInterchangeAnnotation SequenceBranch(
        string id,
        string text,
        SequenceArtifactBlockKind parentKind,
        string branchKind,
        int start,
        int end,
        int depth) => new() {
            Id = id,
            Role = VisualArtifactInterchangeAnnotationRole.SequenceBranch,
            Kind = "SequenceBranch:" + branchKind,
            Text = text,
            StartIndex = start,
            EndIndex = end,
            Sequence = new VisualArtifactInterchangeSequenceAnnotation {
                ParentBlockKind = parentKind,
                BranchKind = branchKind,
                Depth = depth
            }
        };

    private static VisualArtifactInterchangeAnnotation SequenceActivation(string id, string participantId, bool active, int row) {
        var annotation = new VisualArtifactInterchangeAnnotation {
            Id = id,
            Role = VisualArtifactInterchangeAnnotationRole.SequenceActivation,
            Kind = active ? "SequenceActivation" : "SequenceDeactivation",
            StartIndex = row,
            EndIndex = row,
            Sequence = new VisualArtifactInterchangeSequenceAnnotation { ActivationState = active }
        };
        annotation.TargetIds.Add(participantId);
        return annotation;
    }

    private static VisualArtifactInterchangeEdge Message(
        string id,
        string source,
        string target,
        string label,
        int order,
        bool activates = false,
        bool deactivates = false,
        bool dashed = false,
        SequenceArtifactMessageKind kind = SequenceArtifactMessageKind.Call) {
        return new VisualArtifactInterchangeEdge {
            Id = id,
            Role = VisualArtifactInterchangeEdgeRole.SequenceMessage,
            Kind = "SequenceMessage",
            SourceId = source,
            TargetId = target,
            Label = label,
            Order = order,
            Sequence = new VisualArtifactInterchangeSequenceEdge {
                Kind = kind,
                LineStyle = dashed ? SequenceArtifactMessageLineStyle.Dashed : SequenceArtifactMessageLineStyle.Solid,
                ActivatesTarget = activates,
                Deactivates = deactivates
            }
        };
    }
}
