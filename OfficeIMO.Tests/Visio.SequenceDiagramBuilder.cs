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

            Assert.Contains("already exists", participantCollision.Message);
            Assert.Contains("already exists", messageCollision.Message);
            Assert.Contains("already exists", noteCollision.Message);
            Assert.Contains("already exists", activationCollision.Message);
            Assert.Contains("already exists", fragmentCollision.Message);
            Assert.Contains("already exists", fragmentLabelCollision.Message);
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

            Assert.Contains("Unknown sequence participant id", unknownParticipant.Message);
            Assert.Contains("zero or greater", badStart.Message);
            Assert.Contains("greater than or equal", badEnd.Message);
            Assert.Equal("participantIds", nullParticipants.ParamName);
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
    }
}
