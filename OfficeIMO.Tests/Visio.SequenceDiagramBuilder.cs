using System;
using System.IO;
using System.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;
using OfficeIMO.Visio.Stencils;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioSequenceDiagramBuilderTests {
        [Fact]
        public void SequenceDiagramBuilderCreatesParticipantsLifelinesAndMessages() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .SequenceDiagram("Checkout Sequence", sequence => sequence
                    .Title()
                    .Theme(VisioStyleTheme.Fluent())
                    .PageSize(6, 4)
                    .Actor("customer", "Customer")
                    .Participant("web", "Web App")
                    .Control("api", "Orders API")
                    .Database("db", "Orders DB")
                    .Call("customer", "web", "Checkout", "checkout")
                    .Call("web", "api", "POST /orders", "post-order")
                    .Async("api", "db", "Persist order", "persist")
                    .Return("api", "web", "201 Created", "created")
                    .SelfMessage("web", "Render receipt", id: "render")
                    .Activation("web", 1, 4, "web-active")
                    .Activation("api", 2, 3, "api-active")
                    .Fragment("alt retry window", 1, 3, new[] { "web", "api", "db" }, "retry-fragment")
                    .Note("api", "Retry path observed", 2, VisioSide.Right, "retry-note"));

            VisioPage page = Assert.Single(document.Pages);
            Assert.Equal("Checkout Sequence", page.Name);
            Assert.True(page.Width >= 6);
            Assert.True(page.Height >= 4);

            VisioShape customer = Assert.Single(page.Shapes, shape => shape.Id == "customer");
            VisioShape web = Assert.Single(page.Shapes, shape => shape.Id == "web");
            VisioShape api = Assert.Single(page.Shapes, shape => shape.Id == "api");
            VisioShape db = Assert.Single(page.Shapes, shape => shape.Id == "db");

            Assert.Equal("Circle", customer.MasterNameU);
            Assert.Equal("Rectangle", web.MasterNameU);
            Assert.Equal("Rectangle", api.MasterNameU);
            Assert.Equal("Data", db.MasterNameU);
            Assert.True(customer.PinX < web.PinX);
            Assert.True(web.PinX < api.PinX);
            Assert.True(api.PinX < db.PinX);
            Assert.Equal("SequenceParticipant", customer.GetUserCellValue("OfficeIMO.Kind"));
            Assert.Equal("Database", db.GetUserCellValue("OfficeIMO.SequenceParticipantKind"));
            VisioShape note = Assert.Single(page.Shapes, shape => shape.Id == "retry-note");
            Assert.Equal("Rectangle", note.MasterNameU);
            Assert.Equal("SequenceNote", note.GetUserCellValue("OfficeIMO.Kind"));
            Assert.Equal("api", note.GetUserCellValue("OfficeIMO.SequenceParticipantId"));
            Assert.Contains("Sequence Notes", note.LayerNames);
            VisioShape webActivation = Assert.Single(page.Shapes, shape => shape.Id == "web-active");
            Assert.Equal("Rectangle", webActivation.MasterNameU);
            Assert.Equal("SequenceActivation", webActivation.GetUserCellValue("OfficeIMO.Kind"));
            Assert.Equal("web", webActivation.GetUserCellValue("OfficeIMO.SequenceParticipantId"));
            Assert.Equal("1", webActivation.GetUserCellValue("OfficeIMO.SequenceStartRowIndex"));
            Assert.Equal("4", webActivation.GetUserCellValue("OfficeIMO.SequenceEndRowIndex"));
            Assert.Contains("Sequence Activations", webActivation.LayerNames);
            Assert.True(webActivation.Height > webActivation.Width);
            VisioShape retryFragment = Assert.Single(page.Shapes, shape => shape.Id == "retry-fragment");
            Assert.Equal("Rectangle", retryFragment.MasterNameU);
            Assert.Equal("SequenceFragment", retryFragment.GetUserCellValue("OfficeIMO.Kind"));
            Assert.Equal("1", retryFragment.GetUserCellValue("OfficeIMO.SequenceStartRowIndex"));
            Assert.Equal("3", retryFragment.GetUserCellValue("OfficeIMO.SequenceEndRowIndex"));
            Assert.Equal("web;api;db", retryFragment.GetUserCellValue("OfficeIMO.SequenceParticipantIds"));
            Assert.Equal(0, retryFragment.FillPattern);
            Assert.Contains("Sequence Fragments", retryFragment.LayerNames);
            Assert.True(retryFragment.PinX < db.PinX);
            Assert.True(retryFragment.Width > db.PinX - web.PinX);
            VisioShape retryFragmentLabel = Assert.Single(page.Shapes, shape => shape.Id == "retry-fragment-label");
            Assert.Equal("alt retry window", retryFragmentLabel.Text);
            Assert.Equal("DiagramAdornment", retryFragmentLabel.GetUserCellValue("OfficeIMO.Kind"));
            Assert.Equal("retry-fragment", retryFragmentLabel.GetUserCellValue("OfficeIMO.SequenceFragmentId"));
            Assert.Contains("Sequence Fragments", retryFragmentLabel.LayerNames);
            VisioStencilProfile profile = document.CreateStencilProfile();
            Assert.Equal(8, profile.StencilBackedShapeCount);
            Assert.Equal(new[] { "Sequence Diagram" }, profile.StencilCatalogs);
            Assert.Contains("SequenceActivation", profile.SemanticKinds);
            Assert.Contains("SequenceFragment", profile.SemanticKinds);
            Assert.Contains(profile.Usages, usage => usage.StencilId == "seq.actor" && usage.Count == 1);
            Assert.Contains(profile.Usages, usage => usage.StencilId == "seq.participant" && usage.Count == 1);
            Assert.Contains(profile.Usages, usage => usage.StencilId == "seq.control" && usage.Count == 1);
            Assert.Contains(profile.Usages, usage => usage.StencilId == "seq.database" && usage.Count == 1);
            Assert.Contains(profile.Usages, usage => usage.StencilId == "seq.activation" && usage.Count == 2);
            Assert.Contains(profile.Usages, usage => usage.StencilId == "seq.fragment" && usage.Count == 1);
            Assert.Contains(profile.Usages, usage => usage.StencilId == "seq.note" && usage.Count == 1);

            VisioConnector[] messageConnectors = page.Connectors
                .Where(connector => !string.IsNullOrWhiteSpace(connector.Label))
                .ToArray();
            Assert.Equal(5, messageConnectors.Length);
            Assert.Contains(messageConnectors, connector => connector.Id == "created" && connector.LinePattern == 2);
            Assert.Contains(messageConnectors, connector => connector.Id == "persist" && connector.EndArrow == EndArrow.Arrow);
            VisioConnector selfMessage = Assert.Single(messageConnectors, connector => connector.Id == "render");
            Assert.Equal(2, selfMessage.Waypoints.Count);
            Assert.NotNull(selfMessage.LabelPlacement);
            Assert.True(selfMessage.LabelPlacement.PinX > selfMessage.Waypoints.Max(waypoint => waypoint.X));
            Assert.Contains(page.Layers, layer => layer.Name == "Sequence Lifelines");
            Assert.Contains(page.Layers, layer => layer.Name == "Sequence Messages");
            Assert.Contains(page.Layers, layer => layer.Name == "Sequence Activations");
            Assert.Contains(page.Layers, layer => layer.Name == "Sequence Fragments");
            Assert.DoesNotContain(page.AnalyzeVisualQuality(), issue =>
                issue.ShapeId == "web-active" ||
                issue.ShapeId == "api-active" ||
                issue.ShapeId == "retry-fragment" ||
                issue.ShapeId == "retry-fragment-label");

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));

            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Contains(loaded.Pages[0].Connectors, connector => connector.Label == "POST /orders");
            Assert.Contains(loaded.Pages[0].Connectors, connector => connector.Label == "Render receipt" && connector.Waypoints.Count == 2);
        }

        [Fact]
        public void SequenceSelfMessageLabelStaysOutsideLoopAndFlipsNearRightEdge() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .SequenceDiagram("Self Message Placement", sequence => sequence
                    .PageSize(5, 4)
                    .Margins(0.65, 0.65, 0.65, 0.65)
                    .ParticipantSize(1, 0.5)
                    .Spacing(0.8, 0.62, 0.5)
                    .Participant("left", "Left")
                    .Participant("right", "Right")
                    .SelfMessage("left", "Short", id: "left-loop")
                    .SelfMessage("right", "Update incident record", id: "right-loop"));

            VisioPage page = Assert.Single(document.Pages);
            VisioConnector leftLoop = Assert.Single(page.Connectors, connector => connector.Id == "left-loop");
            VisioConnector rightLoop = Assert.Single(page.Connectors, connector => connector.Id == "right-loop");

            Assert.NotNull(leftLoop.LabelPlacement);
            Assert.NotNull(rightLoop.LabelPlacement);
            Assert.True(leftLoop.LabelPlacement.PinX > leftLoop.Waypoints.Max(waypoint => waypoint.X));
            Assert.True(rightLoop.LabelPlacement.PinX < rightLoop.Waypoints.Min(waypoint => waypoint.X));
            Assert.InRange(rightLoop.LabelPlacement.Width, 0.9D, 2.4D);

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void SequenceMessageLabelsPreferLifelineGapsForLongSpanningMessages() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .SequenceDiagram("Lifeline Label Placement", sequence => sequence
                    .PageSize(8, 4.8)
                    .Margins(0.65, 0.65, 0.65, 0.65)
                    .ParticipantSize(1.0, 0.5)
                    .Spacing(1.5, 0.72, 0.5)
                    .Participant("client", "Client")
                    .Participant("api", "API")
                    .Participant("queue", "Queue")
                    .Activation("api", 0, 1, "api-active")
                    .Call("client", "queue", "Submit order for asynchronous processing", "submit"));

            VisioPage page = Assert.Single(document.Pages);
            VisioShape api = Assert.Single(page.Shapes, shape => shape.Id == "api");
            VisioShape queue = Assert.Single(page.Shapes, shape => shape.Id == "queue");
            VisioConnector submit = Assert.Single(page.Connectors, connector => connector.Id == "submit");

            Assert.NotNull(submit.LabelPlacement);
            double labelLeft = submit.LabelPlacement!.PinX!.Value - (submit.LabelPlacement.Width / 2D);
            double labelRight = submit.LabelPlacement.PinX.Value + (submit.LabelPlacement.Width / 2D);
            Assert.NotInRange(submit.LabelPlacement.PinX.Value, api.PinX - 0.08D, api.PinX + 0.08D);
            Assert.True(labelRight < api.PinX - 0.08D || labelLeft > api.PinX + 0.08D);
            Assert.True(labelRight < queue.PinX - 0.08D);

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void SequenceMessageLabelsAvoidFragmentGuardLabels() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .SequenceDiagram("Fragment Label Placement", sequence => sequence
                    .PageSize(8, 4.8)
                    .Margins(0.65, 0.65, 0.65, 0.65)
                    .ParticipantSize(1.0, 0.5)
                    .Spacing(1.5, 0.72, 0.5)
                    .Actor("support", "Support")
                    .Participant("monitor", "Monitor")
                    .Control("api", "API")
                    .Call("support", "api", "Gateway latency", "gateway-latency")
                    .Fragment("alt recovery", 0, 2, new[] { "support", "monitor", "api" }, "recovery")
                    .FragmentGuard("recovery", "[timeout elevated]", 0, "timeout-guard"));

            VisioPage page = Assert.Single(document.Pages);
            VisioShape guardLabel = Assert.Single(page.Shapes, shape => shape.Id == "timeout-guard-label");
            VisioConnector gatewayLatency = Assert.Single(page.Connectors, connector => connector.Id == "gateway-latency");

            Assert.NotNull(gatewayLatency.LabelPlacement);
            Assert.False(ConnectorLabelNearShape(gatewayLatency, guardLabel, 0.06D));

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void SequenceNotesAvoidConnectorLabelsAndStackOnCrowdedRows() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .SequenceDiagram("Crowded Notes", sequence => sequence
                    .PageSize(9, 5)
                    .Margins(0.65, 0.65, 0.65, 0.65)
                    .ParticipantSize(1, 0.5)
                    .Spacing(2.0, 0.82, 0.5)
                    .Participant("web", "Web")
                    .Participant("api", "API")
                    .Participant("db", "DB")
                    .Call("web", "api", "POST /orders", "post-order")
                    .Call("api", "db", "Persist", "persist")
                    .Return("db", "api", "Saved", "saved")
                    .Note("api", "Retry starts here", 1, VisioSide.Right, "retry-note")
                    .Note("api", "Escalate if repeated", 1, VisioSide.Right, "escalation-note"));

            VisioPage page = Assert.Single(document.Pages);
            VisioShape retryNote = Assert.Single(page.Shapes, shape => shape.Id == "retry-note");
            VisioShape escalationNote = Assert.Single(page.Shapes, shape => shape.Id == "escalation-note");

            Assert.Equal("Right", retryNote.GetUserCellValue("OfficeIMO.SequenceRequestedPlacement"));
            Assert.False(ShapesOverlap(retryNote, escalationNote));
            Assert.True(Math.Abs(retryNote.PinY - escalationNote.PinY) > 0.2D);
            Assert.DoesNotContain(page.AnalyzeVisualQuality(), issue =>
                issue.ShapeId == "retry-note" ||
                issue.ShapeId == "escalation-note" ||
                issue.OtherShapeId == "retry-note" ||
                issue.OtherShapeId == "escalation-note");

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void SequenceFragmentPartitionsRenderGuardLabelsAndMetadata() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .SequenceDiagram("Partitioned Fragment", sequence => sequence
                    .PageSize(8, 5)
                    .Margins(0.65, 0.65, 0.65, 0.65)
                    .ParticipantSize(1, 0.5)
                    .Spacing(1.4, 0.72, 0.5)
                    .Actor("support", "Support")
                    .Control("api", "API")
                    .Database("ledger", "Ledger")
                    .Call("support", "api", "Alert", "alert")
                    .Call("support", "api", "Check health", "check")
                    .Call("api", "ledger", "Verify settlement", "verify")
                    .Return("ledger", "api", "Consistent", "consistent")
                    .Async("api", "support", "Resume drain", "resume")
                    .Fragment("alt recovery", 1, 4, new[] { "support", "api", "ledger" }, "recovery")
                    .FragmentGuard("recovery", "[timeout elevated]", 1, "timeout-guard")
                    .FragmentPartition("recovery", "[settlement confirmed]", 3, "settlement-operand"));

            VisioPage page = Assert.Single(document.Pages);
            VisioShape guardLabel = Assert.Single(page.Shapes, shape => shape.Id == "timeout-guard-label");
            VisioShape partitionLabel = Assert.Single(page.Shapes, shape => shape.Id == "settlement-operand-label");
            VisioConnector partitionDivider = Assert.Single(page.Connectors, connector => connector.Id == "settlement-operand");

            Assert.Equal("[timeout elevated]", guardLabel.Text);
            Assert.Equal("[settlement confirmed]", partitionLabel.Text);
            Assert.Equal("DiagramAdornment", guardLabel.GetUserCellValue("OfficeIMO.Kind"));
            Assert.Equal("recovery", guardLabel.GetUserCellValue("OfficeIMO.SequenceFragmentId"));
            Assert.Equal("timeout-guard", guardLabel.GetUserCellValue("OfficeIMO.SequenceFragmentOperandId"));
            Assert.Equal("1", guardLabel.GetUserCellValue("OfficeIMO.SequenceFragmentOperandRowIndex"));
            Assert.Equal("false", guardLabel.GetUserCellValue("OfficeIMO.SequenceFragmentOperandDivider"));
            Assert.Equal("settlement-operand", partitionLabel.GetUserCellValue("OfficeIMO.SequenceFragmentOperandId"));
            Assert.Equal("3", partitionLabel.GetUserCellValue("OfficeIMO.SequenceFragmentOperandRowIndex"));
            Assert.Equal("true", partitionLabel.GetUserCellValue("OfficeIMO.SequenceFragmentOperandDivider"));
            Assert.Equal(2, partitionDivider.LinePattern);
            Assert.Equal(EndArrow.None, partitionDivider.EndArrow);
            Assert.True(partitionDivider.From.PinY > partitionLabel.PinY);
            Assert.Contains("Sequence Fragments", guardLabel.LayerNames);
            Assert.Contains("Sequence Fragments", partitionLabel.LayerNames);

            Assert.DoesNotContain(page.AnalyzeVisualQuality(), issue =>
                issue.ShapeId == "timeout-guard-label" ||
                issue.ShapeId == "settlement-operand-label" ||
                issue.OtherShapeId == "timeout-guard-label" ||
                issue.OtherShapeId == "settlement-operand-label");

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void SequenceNestedFragmentsRenderInsideParentWithOverlapLanes() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .SequenceDiagram("Nested Fragments", sequence => sequence
                    .PageSize(9.5, 6)
                    .Margins(0.75, 0.65, 0.75, 0.65)
                    .ParticipantSize(1, 0.5)
                    .Spacing(1.35, 0.7, 0.5)
                    .Actor("support", "Support")
                    .Control("api", "API")
                    .Database("ledger", "Ledger")
                    .Entity("runbook", "Runbook")
                    .Call("support", "api", "Alert", "alert")
                    .Call("api", "ledger", "Check state", "check")
                    .Return("ledger", "api", "Timeout", "timeout")
                    .Async("support", "runbook", "Open recovery", "open")
                    .Call("api", "ledger", "Retry payment", "retry")
                    .Return("ledger", "api", "Accepted", "accepted")
                    .Fragment("alt incident recovery", 1, 5, new[] { "support", "api", "ledger", "runbook" }, "recovery")
                    .NestedFragment("recovery", "opt retry payment", 2, 4, new[] { "api", "ledger" }, "retry-fragment")
                    .NestedFragment("recovery", "par operator evidence", 2, 4, new[] { "support", "api" }, "evidence-fragment")
                    .FragmentGuard("retry-fragment", "[transient fault]", 2, "retry-guard")
                    .FragmentGuard("evidence-fragment", "[manual evidence]", 3, "evidence-guard"));

            VisioPage page = Assert.Single(document.Pages);
            VisioShape parent = Assert.Single(page.Shapes, shape => shape.Id == "recovery");
            VisioShape retry = Assert.Single(page.Shapes, shape => shape.Id == "retry-fragment");
            VisioShape evidence = Assert.Single(page.Shapes, shape => shape.Id == "evidence-fragment");

            Assert.Equal("recovery", retry.GetUserCellValue("OfficeIMO.SequenceParentFragmentId"));
            Assert.Equal("1", retry.GetUserCellValue("OfficeIMO.SequenceFragmentDepth"));
            Assert.Equal("0", retry.GetUserCellValue("OfficeIMO.SequenceFragmentOverlapLane"));
            Assert.Equal("recovery", evidence.GetUserCellValue("OfficeIMO.SequenceParentFragmentId"));
            Assert.Equal("1", evidence.GetUserCellValue("OfficeIMO.SequenceFragmentDepth"));
            Assert.Equal("1", evidence.GetUserCellValue("OfficeIMO.SequenceFragmentOverlapLane"));
            Assert.True(Left(retry) > Left(parent));
            Assert.True(Right(retry) < Right(parent));
            Assert.True(Top(retry) < Top(parent));
            Assert.True(Bottom(retry) > Bottom(parent));
            Assert.True(evidence.Width < parent.Width);
            Assert.Contains(page.Shapes, shape => shape.Id == "retry-guard-label" && shape.Text == "[transient fault]");
            Assert.Contains(page.Shapes, shape => shape.Id == "evidence-guard-label" && shape.Text == "[manual evidence]");

            Assert.DoesNotContain(page.AnalyzeVisualQuality(), issue =>
                issue.ShapeId == "retry-fragment-label" ||
                issue.ShapeId == "evidence-fragment-label" ||
                issue.OtherShapeId == "retry-fragment-label" ||
                issue.OtherShapeId == "evidence-fragment-label");

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void SequenceDiagramBuilderImportsRecordSets() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioSequenceParticipantRecord[] participants = {
                new("support", "Support", VisioSequenceParticipantKind.Actor),
                new("api", "API", VisioSequenceParticipantKind.Control),
                new("runbook", "Runbook", VisioSequenceParticipantKind.Entity)
            };
            VisioSequenceMessageRecord[] messages = {
                new("alert", "api", "support", "Alert"),
                new("open", "support", "runbook", "Open runbook"),
                new("record", "support", "support", "Update record", selfMessage: true)
            };
            VisioSequenceActivationRecord[] activations = {
                new("support-active", "support", 0, 2)
            };
            VisioSequenceFragmentRecord[] fragments = {
                new("recovery", "alt recovery", 0, 2, new[] { "support", "api", "runbook" }),
                new("runbook-check", "opt runbook evidence", 1, 2, new[] { "support", "runbook" }, "recovery")
            };
            VisioSequenceFragmentOperandRecord[] operands = {
                new("guard", "recovery", "[active incident]", 0)
            };
            VisioSequenceNoteRecord[] notes = {
                new("runbook-note", "runbook", "Checklist", 1, VisioSide.Left)
            };

            VisioDocument document = VisioDocument.Create(filePath)
                .SequenceDiagram("Imported Incident", sequence => sequence
                    .PageSize(7, 4.8)
                    .Import(participants, messages, activations, fragments, operands, notes));

            VisioPage page = Assert.Single(document.Pages);
            Assert.Contains(page.Shapes, shape => shape.Id == "support" && shape.GetUserCellValue("OfficeIMO.SequenceParticipantKind") == "Actor");
            Assert.Contains(page.Shapes, shape => shape.Id == "support-active" && shape.GetUserCellValue("OfficeIMO.Kind") == "SequenceActivation");
            Assert.Contains(page.Shapes, shape => shape.Id == "recovery" && shape.GetUserCellValue("OfficeIMO.SequenceParticipantIds") == "support;api;runbook");
            Assert.Contains(page.Shapes, shape => shape.Id == "runbook-check" && shape.GetUserCellValue("OfficeIMO.SequenceParentFragmentId") == "recovery");
            Assert.Contains(page.Shapes, shape => shape.Id == "guard-label" && shape.Text == "[active incident]");
            Assert.Contains(page.Shapes, shape => shape.Id == "runbook-note" && shape.GetUserCellValue("OfficeIMO.SequenceRequestedPlacement") == "Left");
            Assert.Contains(page.Connectors, connector => connector.Id == "record" && connector.Waypoints.Count == 2);

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void SequenceDiagramBuilderRejectsUnknownParticipants() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentException exception = Assert.Throws<ArgumentException>(() =>
                document.SequenceDiagram("Invalid", sequence => sequence
                    .Participant("web", "Web")
                    .Call("web", "missing", "request")));

            Assert.Contains("Unknown sequence participant id", exception.Message);
        }

        [Fact]
        public void SequenceDiagramBuilderRejectsDuplicateIds() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentException participantCollision = Assert.Throws<ArgumentException>(() =>
                document.SequenceDiagram("Invalid", sequence => sequence
                    .Participant("web", "Web")
                    .Participant("web", "Duplicate")));
            ArgumentException messageCollision = Assert.Throws<ArgumentException>(() =>
                document.SequenceDiagram("Invalid", sequence => sequence
                    .Participant("web", "Web")
                    .Participant("api", "API")
                    .Call("web", "api", "first", "same")
                    .Return("api", "web", "second", "same")));
            ArgumentException noteCollision = Assert.Throws<ArgumentException>(() =>
                document.SequenceDiagram("Invalid", sequence => sequence
                    .Participant("web", "Web")
                    .Note("web", "First", 0, id: "same")
                    .Call("web", "web", "Duplicate", "same")));
            ArgumentException activationCollision = Assert.Throws<ArgumentException>(() =>
                document.SequenceDiagram("Invalid", sequence => sequence
                    .Participant("web", "Web")
                    .Activation("web", 0, 1, "same")
                    .Note("web", "Duplicate", 0, id: "same")));
            ArgumentException fragmentCollision = Assert.Throws<ArgumentException>(() =>
                document.SequenceDiagram("Invalid", sequence => sequence
                    .Participant("web", "Web")
                    .Fragment("opt", 0, 1, "same")
                    .Call("web", "web", "Duplicate", "same")));
            ArgumentException fragmentLabelCollision = Assert.Throws<ArgumentException>(() =>
                document.SequenceDiagram("Invalid", sequence => sequence
                    .Participant("web", "Web")
                    .Fragment("opt", 0, 1, "same")
                    .Note("web", "Duplicate", 0, id: "same-label")));
            ArgumentException fragmentOperandLabelCollision = Assert.Throws<ArgumentException>(() =>
                document.SequenceDiagram("Invalid", sequence => sequence
                    .Participant("web", "Web")
                    .Fragment("opt", 0, 1, "fragment")
                    .FragmentPartition("fragment", "[else]", 1, "same")
                    .Note("web", "Duplicate", 0, id: "same-label")));

            Assert.Contains("already exists", participantCollision.Message);
            Assert.Contains("already exists", messageCollision.Message);
            Assert.Contains("already exists", noteCollision.Message);
            Assert.Contains("already exists", activationCollision.Message);
            Assert.Contains("already exists", fragmentCollision.Message);
            Assert.Contains("already exists", fragmentLabelCollision.Message);
            Assert.Contains("already exists", fragmentOperandLabelCollision.Message);
        }

        [Fact]
        public void SequenceDiagramBuilderRejectsInvalidNotes() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentException unknownTarget = Assert.Throws<ArgumentException>(() =>
                document.SequenceDiagram("Invalid", sequence => sequence
                    .Participant("web", "Web")
                    .Note("missing", "No target", 0)));
            ArgumentOutOfRangeException badRow = Assert.Throws<ArgumentOutOfRangeException>(() =>
                document.SequenceDiagram("Invalid", sequence => sequence
                    .Participant("web", "Web")
                    .Note("web", "Bad row", -1)));
            ArgumentOutOfRangeException badPlacement = Assert.Throws<ArgumentOutOfRangeException>(() =>
                document.SequenceDiagram("Invalid", sequence => sequence
                    .Participant("web", "Web")
                    .Note("web", "Bad placement", 0, VisioSide.Top)));

            Assert.Contains("Unknown sequence participant id", unknownTarget.Message);
            Assert.Contains("zero or greater", badRow.Message);
            Assert.Contains("left or right", badPlacement.Message);
        }

        [Fact]
        public void SequenceDiagramBuilderRejectsInvalidActivations() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentException unknownTarget = Assert.Throws<ArgumentException>(() =>
                document.SequenceDiagram("Invalid", sequence => sequence
                    .Participant("web", "Web")
                    .Activation("missing", 0, 1)));
            ArgumentOutOfRangeException badStart = Assert.Throws<ArgumentOutOfRangeException>(() =>
                document.SequenceDiagram("Invalid", sequence => sequence
                    .Participant("web", "Web")
                    .Activation("web", -1, 1)));
            ArgumentOutOfRangeException badEnd = Assert.Throws<ArgumentOutOfRangeException>(() =>
                document.SequenceDiagram("Invalid", sequence => sequence
                    .Participant("web", "Web")
                    .Activation("web", 2, 1)));

            Assert.Contains("Unknown sequence participant id", unknownTarget.Message);
            Assert.Contains("zero or greater", badStart.Message);
            Assert.Contains("greater than or equal", badEnd.Message);
        }

        [Fact]
        public void SequenceDiagramBuilderRejectsInvalidFragments() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentException unknownParticipant = Assert.Throws<ArgumentException>(() =>
                document.SequenceDiagram("Invalid", sequence => sequence
                    .Participant("web", "Web")
                    .Fragment("alt", 0, 1, new[] { "missing" })));
            ArgumentOutOfRangeException badStart = Assert.Throws<ArgumentOutOfRangeException>(() =>
                document.SequenceDiagram("Invalid", sequence => sequence
                    .Participant("web", "Web")
                    .Fragment("alt", -1, 1)));
            ArgumentOutOfRangeException badEnd = Assert.Throws<ArgumentOutOfRangeException>(() =>
                document.SequenceDiagram("Invalid", sequence => sequence
                    .Participant("web", "Web")
                    .Fragment("alt", 2, 1)));
            ArgumentNullException nullParticipants = Assert.Throws<ArgumentNullException>(() =>
                document.SequenceDiagram("Invalid", sequence => sequence
                    .Participant("web", "Web")
                    .Fragment("alt", 0, 1, participantIds: null!)));
            ArgumentException unknownParent = Assert.Throws<ArgumentException>(() =>
                document.SequenceDiagram("Invalid", sequence => sequence
                    .Participant("web", "Web")
                    .NestedFragment("missing", "opt", 0, 1, "nested")));
            ArgumentOutOfRangeException nestedOutsideParent = Assert.Throws<ArgumentOutOfRangeException>(() =>
                document.SequenceDiagram("Invalid", sequence => sequence
                    .Participant("web", "Web")
                    .Fragment("alt", 1, 2, "parent")
                    .NestedFragment("parent", "opt", 0, 1, "nested")));
            ArgumentException nestedParticipantOutsideParent = Assert.Throws<ArgumentException>(() =>
                document.SequenceDiagram("Invalid", sequence => sequence
                    .Participant("web", "Web")
                    .Participant("api", "API")
                    .Fragment("alt", 0, 2, new[] { "web" }, "parent")
                    .NestedFragment("parent", "opt", 1, 2, new[] { "api" }, "nested")));

            Assert.Contains("Unknown sequence participant id", unknownParticipant.Message);
            Assert.Contains("zero or greater", badStart.Message);
            Assert.Contains("greater than or equal", badEnd.Message);
            Assert.Equal("participantIds", nullParticipants.ParamName);
            Assert.Contains("Unknown sequence fragment id", unknownParent.Message);
            Assert.Contains("inside the parent fragment row range", nestedOutsideParent.Message);
            Assert.Contains("outside parent fragment", nestedParticipantOutsideParent.Message);
        }

        [Fact]
        public void SequenceDiagramBuilderRejectsInvalidFragmentOperands() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentException unknownFragment = Assert.Throws<ArgumentException>(() =>
                document.SequenceDiagram("Invalid", sequence => sequence
                    .Participant("web", "Web")
                    .FragmentGuard("missing", "[guard]", 0)));
            ArgumentOutOfRangeException guardOutside = Assert.Throws<ArgumentOutOfRangeException>(() =>
                document.SequenceDiagram("Invalid", sequence => sequence
                    .Participant("web", "Web")
                    .Fragment("opt", 1, 2, "fragment")
                    .FragmentGuard("fragment", "[guard]", 0)));
            ArgumentOutOfRangeException partitionAtFirstRow = Assert.Throws<ArgumentOutOfRangeException>(() =>
                document.SequenceDiagram("Invalid", sequence => sequence
                    .Participant("web", "Web")
                    .Fragment("alt", 1, 2, "fragment")
                    .FragmentPartition("fragment", "[else]", 1)));

            Assert.Contains("Unknown sequence fragment id", unknownFragment.Message);
            Assert.Contains("inside the fragment row range", guardOutside.Message);
            Assert.Contains("after the first fragment row", partitionAtFirstRow.Message);
        }

        [Fact]
        public void SequenceStencilsAreSearchableAndIncludedInAllCatalog() {
            VisioStencilShape participant = VisioStencils.Sequence.Get("seq.participant");
            VisioStencilShape actor = Assert.Single(VisioStencils.Sequence.Search("person"));
            VisioStencilShape activation = Assert.Single(VisioStencils.Sequence.Search("execution"));
            VisioStencilShape fragment = Assert.Single(VisioStencils.Sequence.Search("combined-fragment"));

            Assert.Equal("Rectangle", participant.MasterNameU);
            Assert.Equal("Actor", actor.Name);
            Assert.Equal("Activation", activation.Name);
            Assert.Equal("Combined Fragment", fragment.Name);
            Assert.Contains(VisioStencils.All.Shapes, shape => shape.Id == "seq.database");
            Assert.Contains(VisioStencils.All.Shapes, shape => shape.Id == "seq.activation");
            Assert.Contains(VisioStencils.All.Shapes, shape => shape.Id == "seq.fragment");
        }

        private static bool ShapesOverlap(VisioShape first, VisioShape second) {
            double left = Math.Max(first.PinX - (first.Width / 2D), second.PinX - (second.Width / 2D));
            double right = Math.Min(first.PinX + (first.Width / 2D), second.PinX + (second.Width / 2D));
            double bottom = Math.Max(first.PinY - (first.Height / 2D), second.PinY - (second.Height / 2D));
            double top = Math.Min(first.PinY + (first.Height / 2D), second.PinY + (second.Height / 2D));
            return right > left && top > bottom;
        }

        private static double Left(VisioShape shape) => shape.PinX - (shape.Width / 2D);

        private static double Right(VisioShape shape) => shape.PinX + (shape.Width / 2D);

        private static double Top(VisioShape shape) => shape.PinY + (shape.Height / 2D);

        private static double Bottom(VisioShape shape) => shape.PinY - (shape.Height / 2D);

        private static bool ConnectorLabelNearShape(VisioConnector connector, VisioShape shape, double padding) {
            Assert.NotNull(connector.LabelPlacement);
            Assert.True(connector.LabelPlacement!.PinX.HasValue);
            Assert.True(connector.LabelPlacement.PinY.HasValue);

            double left = Math.Max(connector.LabelPlacement.PinX.Value - (connector.LabelPlacement.Width / 2D), shape.PinX - (shape.Width / 2D) - padding);
            double right = Math.Min(connector.LabelPlacement.PinX.Value + (connector.LabelPlacement.Width / 2D), shape.PinX + (shape.Width / 2D) + padding);
            double bottom = Math.Max(connector.LabelPlacement.PinY.Value - (connector.LabelPlacement.Height / 2D), shape.PinY - (shape.Height / 2D) - padding);
            double top = Math.Min(connector.LabelPlacement.PinY.Value + (connector.LabelPlacement.Height / 2D), shape.PinY + (shape.Height / 2D) + padding);
            return right > left && top > bottom;
        }
    }
}
