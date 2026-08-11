using System.Linq;
using global::ChartForgeX.Primitives;
using global::ChartForgeX.Topology;
using global::ChartForgeX.VisualArtifacts;
using OfficeIMO.ChartForgeX;
using Xunit;

namespace OfficeIMO.ChartForgeX.Tests;

public sealed partial class OfficeVisioVisualIntegrationTests {
    [Fact]
    public void GraphAndFlowNodeDetailsReportTypedSemanticLoss() {
        VisualArtifactInterchangeEnvelope topology = TopologyEnvelope("topology-details");
        VisualArtifactInterchangeNode topologyNode = TopologyNode("service", "Service");
        topologyNode.Details.Add(new VisualArtifactInterchangeDetail { Label = "Region", Value = "EU" });
        topology.Nodes.Add(topologyNode);

        OfficeVisioVisualConversionResult topologyResult = topology.ToOfficeVisio();

        Assert.Contains(topologyResult.Report.Diagnostics, diagnostic =>
            diagnostic.Code == OfficeVisioVisualDiagnosticCode.DetailsNotRendered &&
            diagnostic.EntityKind == OfficeVisioVisualEntityKind.Node &&
            diagnostic.EntityId == "service" &&
            diagnostic.Feature == "details");
        Assert.Equal("EU", topologyResult.Page.Shapes.Single(shape => shape.Id == "service").GetShapeDataValue("Detail.1.Region"));

        var flow = new VisualArtifactInterchangeEnvelope {
            Id = "flow-details",
            Kind = VisualArtifactKind.Flow,
            Family = VisualArtifactInterchangeFamily.Flow,
            Flow = new VisualArtifactInterchangeFlowArtifact {
                LayoutMode = FlowArtifactLayoutMode.Layered,
                LayoutDirection = FlowArtifactDirection.LeftToRight
            }
        };
        var flowNode = new VisualArtifactInterchangeNode {
            Id = "approve",
            Role = VisualArtifactInterchangeNodeRole.FlowStep,
            Kind = FlowArtifactStepKind.Process.ToString(),
            Label = "Approve",
            Flow = new VisualArtifactInterchangeFlowNode { Kind = FlowArtifactStepKind.Process }
        };
        flowNode.Details.Add(new VisualArtifactInterchangeDetail { Label = "Owner", Value = "Finance" });
        flow.Nodes.Add(flowNode);

        OfficeVisioVisualConversionResult flowResult = flow.ToOfficeVisio();

        Assert.Contains(flowResult.Report.Diagnostics, diagnostic =>
            diagnostic.Code == OfficeVisioVisualDiagnosticCode.DetailsNotRendered &&
            diagnostic.EntityKind == OfficeVisioVisualEntityKind.Node &&
            diagnostic.EntityId == "approve" &&
            diagnostic.Feature == "details");
        Assert.Equal("Finance", flowResult.Page.Shapes.Single(shape => shape.Id == "approve").GetShapeDataValue("Detail.1.Owner"));
    }

    [Fact]
    public void SequenceAdornmentsDetailsAndEndpointLabelsReportTypedSemanticLoss() {
        VisualArtifactInterchangeEnvelope envelope = SequenceEnvelope("sequence-fidelity");
        VisualArtifactInterchangeNode client = Participant("client", "Client", SequenceArtifactParticipantKind.Actor, 0);
        client.IconId = "person";
        client.Symbol = "C";
        client.Badge = "External";
        client.Details.Add(new VisualArtifactInterchangeDetail { Label = "Region", Value = "EU" });
        envelope.Nodes.Add(client);
        envelope.Nodes.Add(Participant("service", "Service", SequenceArtifactParticipantKind.Control, 1));
        VisualArtifactInterchangeEdge message = Message("request", "client", "service", "Request", 0);
        message.SourceLabel = "caller";
        message.TargetLabel = "callee";
        envelope.Edges.Add(message);

        OfficeVisioVisualConversionResult result = envelope.ToOfficeVisio();

        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == OfficeVisioVisualDiagnosticCode.ArtworkNotProjected &&
            diagnostic.EntityKind == OfficeVisioVisualEntityKind.Participant &&
            diagnostic.EntityId == "client" &&
            diagnostic.Feature == "participantAdornment");
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == OfficeVisioVisualDiagnosticCode.DetailsNotRendered &&
            diagnostic.EntityKind == OfficeVisioVisualEntityKind.Participant &&
            diagnostic.EntityId == "client" &&
            diagnostic.Feature == "details");
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == OfficeVisioVisualDiagnosticCode.EndpointLabelsNotRendered &&
            diagnostic.EntityKind == OfficeVisioVisualEntityKind.Message &&
            diagnostic.EntityId == "request" &&
            diagnostic.Feature == "endpointLabels");
        Assert.Equal("person", result.Page.Shapes.Single(shape => shape.Id == "client").GetShapeDataValue("CFX.Icon"));
        Assert.Equal("EU", result.Page.Shapes.Single(shape => shape.Id == "client").GetShapeDataValue("Detail.1.Region"));
        Assert.Equal("caller", result.Page.Connectors.Single(connector => connector.Id == "request").GetShapeDataValue("CFX.SourceLabel"));
        Assert.Equal("callee", result.Page.Connectors.Single(connector => connector.Id == "request").GetShapeDataValue("CFX.TargetLabel"));
        Assert.True(result.Report.HasSemanticLoss);
    }
}
