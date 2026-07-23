using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Capabilities;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.Tests {
    public partial class PowerPointLegacyPptInteractionTests {
        [Fact]
        public void NativeWriter_AuthorsAndProjectsMacroAndProgramActions() {
            const string MacroName = "Module1.RunReport";
            var programUri = new Uri("file:///Applications/Calculator.app");
            byte[] bytes;
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                PowerPointSlide slide = source.AddSlide(P.SlideLayoutValues.Blank);
                PowerPointAutoShape macroShape = slide.AddRectangle(
                    100000, 100000, 1000000, 500000);
                GetDrawingProperties(macroShape).Append(new A.HyperlinkOnClick {
                    Id = string.Empty,
                    Action = "ppaction://macro?name=" + MacroName,
                    HighlightClick = true
                });

                PowerPointTextBox textBox = slide.AddTextBox("Run tool",
                    100000, 800000, 1000000, 500000);
                HyperlinkRelationship programRelationship = slide.SlidePart
                    .AddHyperlinkRelationship(programUri, true);
                A.Run run = ((P.Shape)textBox.Element).TextBody!
                    .Descendants<A.Run>().Single();
                run.RunProperties ??= new A.RunProperties();
                run.RunProperties.Append(new A.HyperlinkOnMouseOver {
                    Id = programRelationship.Id,
                    Action = "ppaction://program",
                    EndSound = true
                });

                LegacyPptWritePreflightReport report = source.AnalyzeLegacyPptWrite();
                Assert.True(report.CanWrite,
                    string.Join(Environment.NewLine, report.Findings));
                bytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation legacy = LegacyPptPresentation.Load(bytes);
            LegacyPptInteraction macro = Assert.Single(legacy.Slides[0].Shapes
                .SelectMany(shape => shape.Interactions));
            Assert.Equal(LegacyPptInteractionAction.Macro, macro.Action);
            Assert.Equal(MacroName, macro.Name);
            Assert.True(macro.IsAnimated);
            Assert.False(macro.StopsSound);
            LegacyPptTextInteraction programRange = Assert.Single(legacy.Slides[0]
                .Shapes.SelectMany(shape => shape.TextBody.Interactions));
            LegacyPptInteraction program = programRange.Interaction;
            Assert.Equal(LegacyPptInteractionTrigger.MouseOver, program.Trigger);
            Assert.Equal(LegacyPptInteractionAction.RunProgram, program.Action);
            Assert.Equal(programUri.OriginalString, program.Name);
            Assert.False(program.IsAnimated);
            Assert.True(program.StopsSound);

            using var input = new MemoryStream(bytes, writable: false);
            using PowerPointPresentation projected = PowerPointPresentation.Load(input);
            A.HyperlinkOnClick projectedMacro = projected.Slides[0].SlidePart.Slide!
                .Descendants<A.HyperlinkOnClick>().Single(link =>
                    link.Action?.Value?.StartsWith("ppaction://macro?name=",
                        StringComparison.Ordinal) == true);
            Assert.Equal("ppaction://macro?name=" + MacroName,
                projectedMacro.Action?.Value);
            Assert.True(projectedMacro.HighlightClick?.Value);
            Assert.Equal(string.Empty, projectedMacro.Id?.Value);
            A.HyperlinkOnMouseOver projectedProgram = projected.Slides[0]
                .SlidePart.Slide!.Descendants<A.HyperlinkOnMouseOver>().Single();
            Assert.Equal("ppaction://program", projectedProgram.Action?.Value);
            Assert.True(projectedProgram.EndSound?.Value);
            HyperlinkRelationship projectedRelationship = projected.Slides[0]
                .SlidePart.HyperlinkRelationships.Single(relationship =>
                    relationship.Id == projectedProgram.Id?.Value);
            Assert.Equal(programUri, projectedRelationship.Uri);
            Assert.Empty(projected.ValidateDocument());
            LegacyPptWritePreflightReport projectedPreflight = projected
                .AnalyzeLegacyPptWrite();
            Assert.False(projectedPreflight.CanWrite);
            Assert.Contains(projectedPreflight.Findings, finding =>
                finding.Code == "PPT-WRITE-PRESERVED-RUN-PROGRAM");
            Assert.Throws<NotSupportedException>(() =>
                projected.ToBytes(PowerPointFileFormat.Ppt));
            Assert.Equal(bytes, projected.ToBytes(PowerPointFileFormat.Ppt,
                new PowerPointSaveOptions {
                    LossPolicy = PowerPointConversionLossPolicy.Allow
                }));
        }

        [Fact]
        public void ImportedNamedAction_AddEditAndRemoveAppendsPreservingRecords() {
            byte[] sourceBytes;
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                source.AddSlide(P.SlideLayoutValues.Blank).AddRectangle(
                    100000, 100000, 1000000, 500000);
                sourceBytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation original = LegacyPptPresentation.Load(sourceBytes);

            byte[] addedBytes;
            using (var input = new MemoryStream(sourceBytes, writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation.Load(input)) {
                GetDrawingProperties(Assert.Single(imported.Slides[0].Shapes))
                    .Append(new A.HyperlinkOnClick {
                        Id = string.Empty,
                        Action = "ppaction://macro?name=Module1.First",
                        HighlightClick = true
                    });
                LegacyPptWritePreflightReport report = imported.AnalyzeLegacyPptWrite();
                Assert.True(report.CanWrite,
                    string.Join(Environment.NewLine, report.Findings));
                addedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation added = LegacyPptPresentation.Load(addedBytes);
            LegacyPptInteraction addedAction = Assert.Single(added.Slides[0].Shapes
                .SelectMany(shape => shape.Interactions));
            Assert.Equal(LegacyPptInteractionAction.Macro, addedAction.Action);
            Assert.Equal("Module1.First", addedAction.Name);
            Assert.True(addedAction.IsAnimated);
            Assert.True(added.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));

            var programUri = new Uri("file:///Applications/Preview.app");
            byte[] editedBytes;
            using (var input = new MemoryStream(addedBytes, writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation.Load(input)) {
                PowerPointSlide slide = imported.Slides[0];
                P.NonVisualDrawingProperties properties = GetDrawingProperties(
                    Assert.Single(slide.Shapes));
                properties.RemoveAllChildren<A.HyperlinkOnClick>();
                HyperlinkRelationship relationship = slide.SlidePart
                    .AddHyperlinkRelationship(programUri, true);
                properties.Append(new A.HyperlinkOnClick {
                    Id = relationship.Id,
                    Action = "ppaction://program",
                    EndSound = true
                });
                Assert.True(imported.AnalyzeLegacyPptWrite().CanWrite);
                editedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation edited = LegacyPptPresentation.Load(editedBytes);
            LegacyPptInteraction editedAction = Assert.Single(edited.Slides[0].Shapes
                .SelectMany(shape => shape.Interactions));
            Assert.Equal(LegacyPptInteractionAction.RunProgram,
                editedAction.Action);
            Assert.Equal(programUri.OriginalString, editedAction.Name);
            Assert.True(editedAction.StopsSound);
            Assert.True(edited.Package.DocumentStream.AsSpan(0,
                    added.Package.DocumentStream.Length)
                .SequenceEqual(added.Package.DocumentStream));

            byte[] removedBytes;
            using (var input = new MemoryStream(editedBytes, writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation.Load(input)) {
                GetDrawingProperties(Assert.Single(imported.Slides[0].Shapes))
                    .RemoveAllChildren<A.HyperlinkOnClick>();
                Assert.True(imported.AnalyzeLegacyPptWrite().CanWrite);
                removedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation removed = LegacyPptPresentation.Load(removedBytes);
            Assert.Empty(removed.Slides[0].Shapes
                .SelectMany(shape => shape.Interactions));
            Assert.True(removed.Package.DocumentStream.AsSpan(0,
                    edited.Package.DocumentStream.Length)
                .SequenceEqual(edited.Package.DocumentStream));
        }

        [Fact]
        public void NativeWriter_BlocksNamedActionDataWithoutBinarySemantics() {
            using PowerPointPresentation source = PowerPointPresentation.Create();
            PowerPointAutoShape shape = source.AddSlide(P.SlideLayoutValues.Blank).AddRectangle(
                100000, 100000, 1000000, 500000);
            GetDrawingProperties(shape).Append(new A.HyperlinkOnClick {
                Id = string.Empty,
                Action = "ppaction://macro?name=Module1.RunReport",
                History = true
            });

            LegacyPptWriteFinding finding = Assert.Single(
                source.AnalyzeLegacyPptWrite().Findings,
                item => item.Code == "PPT-WRITE-INTERACTION");
            Assert.Contains("history", finding.Description,
                StringComparison.OrdinalIgnoreCase);
        }
    }
}
