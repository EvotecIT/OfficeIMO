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
                    .SelfMessage("web", "Render receipt", id: "render"));

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

            VisioConnector[] messageConnectors = page.Connectors
                .Where(connector => !string.IsNullOrWhiteSpace(connector.Label))
                .ToArray();
            Assert.Equal(5, messageConnectors.Length);
            Assert.Contains(messageConnectors, connector => connector.Id == "created" && connector.LinePattern == 2);
            Assert.Contains(messageConnectors, connector => connector.Id == "persist" && connector.EndArrow == EndArrow.Arrow);
            VisioConnector selfMessage = Assert.Single(messageConnectors, connector => connector.Id == "render");
            Assert.Equal(2, selfMessage.Waypoints.Count);
            Assert.NotNull(selfMessage.LabelPlacement);
            Assert.Contains(page.Layers, layer => layer.Name == "Sequence Lifelines");
            Assert.Contains(page.Layers, layer => layer.Name == "Sequence Messages");

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));

            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Contains(loaded.Pages[0].Connectors, connector => connector.Label == "POST /orders");
            Assert.Contains(loaded.Pages[0].Connectors, connector => connector.Label == "Render receipt" && connector.Waypoints.Count == 2);
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

            Assert.Contains("already exists", participantCollision.Message);
            Assert.Contains("already exists", messageCollision.Message);
        }

        [Fact]
        public void SequenceStencilsAreSearchableAndIncludedInAllCatalog() {
            VisioStencilShape participant = VisioStencils.Sequence.Get("seq.participant");
            VisioStencilShape actor = Assert.Single(VisioStencils.Sequence.Search("person"));

            Assert.Equal("Rectangle", participant.MasterNameU);
            Assert.Equal("Actor", actor.Name);
            Assert.Contains(VisioStencils.All.Shapes, shape => shape.Id == "seq.database");
        }
    }
}
