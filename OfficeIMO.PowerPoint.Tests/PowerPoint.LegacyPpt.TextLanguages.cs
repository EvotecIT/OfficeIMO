using System.Buffers.Binary;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Diagnostics;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using OfficeIMO.PowerPoint.LegacyPpt.Write;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.Tests {
    public partial class PowerPointLegacyPptTests {
        [Fact]
        public void NeutralReader_DecodesMicrosoftTextSpecialInfoRuns() {
            LegacyPptPresentation legacy = LegacyPptPresentation.Load(
                FixturePath);
            LegacyPptTextBody[] bodies = legacy.Slides
                .SelectMany(slide => slide.Shapes)
                .Select(shape => shape.TextBody)
                .Where(body => body.HasTextSpecialInfoRecord)
                .ToArray();

            Assert.NotEmpty(bodies);
            Assert.DoesNotContain(bodies,
                body => body.IsTextSpecialInfoTruncated);
            Assert.Contains(bodies.SelectMany(body => body.LanguageRuns),
                run => run.Language == "pl-PL");
            Assert.DoesNotContain(legacy.Diagnostics, diagnostic =>
                diagnostic.Code == "PPT-TEXT-SPECIAL-INFO-TRUNCATED");
        }

        [Fact]
        public void NativeWriter_AuthorsProjectsAndPreservesTextLanguages() {
            byte[] bytes;
            using (PowerPointPresentation presentation =
                   PowerPointPresentation.Create()) {
                PowerPointTextBox textBox = presentation.AddSlide(
                        P.SlideLayoutValues.Blank)
                    .AddTextBoxPoints(string.Empty, 30, 30, 360, 120);
                P.Shape shape = Assert.IsType<P.Shape>(textBox.Element);
                shape.TextBody = new P.TextBody(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(
                        new A.Run(
                            new A.RunProperties {
                                Language = "en-US",
                                AlternativeLanguage = "fr-FR"
                            },
                            new A.Text("English ")),
                        new A.Run(
                            new A.RunProperties { Language = "pl-PL" },
                            new A.Text("Polski")),
                        new A.EndParagraphRunProperties {
                            Language = "de-DE"
                        }),
                    new A.Paragraph(
                        new A.Run(
                            new A.RunProperties { NoProof = true },
                            new A.Text("Unproofed")),
                        new A.EndParagraphRunProperties {
                            NoProof = true
                        }));

                LegacyPptWritePreflightReport preflight = presentation
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                bytes = presentation.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptTextBody textBody = Assert.Single(Assert.Single(
                LegacyPptPresentation.Load(bytes).Slides).Shapes).TextBody;
            Assert.Equal("English Polski\nUnproofed", textBody.Text);
            Assert.True(textBody.HasTextSpecialInfoRecord);
            Assert.False(textBody.IsTextSpecialInfoTruncated);
            Assert.False(textBody.HasUnprojectedTextSpecialInfo);
            Assert.Collection(textBody.LanguageRuns,
                run => {
                    Assert.Equal(0, run.Start);
                    Assert.Equal(8, run.Length);
                    Assert.Equal((ushort)0x0409, run.LanguageId);
                    Assert.Equal("en-US", run.Language);
                    Assert.Equal((ushort)0x040C,
                        run.AlternativeLanguageId);
                    Assert.Equal("fr-FR", run.AlternativeLanguage);
                },
                run => {
                    Assert.Equal(8, run.Start);
                    Assert.Equal(6, run.Length);
                    Assert.Equal((ushort)0x0415, run.LanguageId);
                    Assert.Equal("pl-PL", run.Language);
                },
                run => {
                    Assert.Equal(14, run.Start);
                    Assert.Equal(1, run.Length);
                    Assert.Equal((ushort)0x0407, run.LanguageId);
                    Assert.Equal("de-DE", run.Language);
                },
                run => {
                    Assert.Equal(15, run.Start);
                    Assert.Equal(10, run.Length);
                    Assert.Equal((ushort)0x0400, run.LanguageId);
                    Assert.True(run.NoProof);
                });

            using var stream = new MemoryStream(bytes, writable: false);
            using PowerPointPresentation reopened = PowerPointPresentation
                .Load(stream);
            P.Shape projected = Assert.IsType<P.Shape>(Assert.Single(
                reopened.Slides[0].TextBoxes).Element);
            A.Paragraph[] paragraphs = projected.TextBody!
                .Elements<A.Paragraph>().ToArray();
            A.Run[] firstRuns = paragraphs[0].Elements<A.Run>().ToArray();
            Assert.Equal("en-US",
                firstRuns[0].RunProperties!.Language!.Value);
            Assert.Equal("fr-FR", firstRuns[0].RunProperties!
                .AlternativeLanguage!.Value);
            Assert.Equal("pl-PL",
                firstRuns[1].RunProperties!.Language!.Value);
            Assert.Equal("de-DE", paragraphs[0]
                .GetFirstChild<A.EndParagraphRunProperties>()!
                .Language!.Value);
            Assert.True(paragraphs[1].Elements<A.Run>().Single()
                .RunProperties!.NoProof!.Value);
            Assert.True(paragraphs[1]
                .GetFirstChild<A.EndParagraphRunProperties>()!
                .NoProof!.Value);
            Assert.Empty(reopened.ValidateDocument());
            Assert.Equal(bytes,
                reopened.ToBytes(PowerPointFileFormat.Ppt));
        }

        [Fact]
        public void NativeWriter_PreservesTrailingEmptyParagraphLanguage() {
            byte[] bytes;
            using (PowerPointPresentation presentation =
                       PowerPointPresentation.Create()) {
                PowerPointTextBox textBox = presentation.AddSlide(
                        P.SlideLayoutValues.Blank)
                    .AddTextBoxPoints("A", 30, 30, 240, 60);
                P.Shape shape = Assert.IsType<P.Shape>(textBox.Element);
                shape.TextBody = new P.TextBody(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(
                        new A.Run(
                            new A.RunProperties { Language = "en-US" },
                            new A.Text("A")),
                        new A.EndParagraphRunProperties {
                            Language = "en-US"
                        }),
                    new A.Paragraph(
                        new A.EndParagraphRunProperties {
                            Language = "pl-PL"
                        }));

                LegacyPptWritePreflightReport preflight = presentation
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                bytes = presentation.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptTextBody binary = Assert.Single(Assert.Single(
                LegacyPptPresentation.Load(bytes).Slides).Shapes).TextBody;
            Assert.Equal("A\n", binary.Text);
            Assert.False(binary.IsTextSpecialInfoTruncated);
            Assert.Collection(binary.LanguageRuns,
                run => {
                    Assert.Equal(0, run.Start);
                    Assert.Equal(2, run.Length);
                    Assert.Equal("en-US", run.Language);
                },
                run => {
                    Assert.Equal(2, run.Start);
                    Assert.Equal(1, run.Length);
                    Assert.Equal("pl-PL", run.Language);
                });

            using var input = new MemoryStream(bytes, writable: false);
            using PowerPointPresentation reopened = PowerPointPresentation
                .Load(input);
            A.Paragraph[] paragraphs = Assert.IsType<P.Shape>(Assert.Single(
                    reopened.Slides[0].TextBoxes).Element).TextBody!
                .Elements<A.Paragraph>().ToArray();
            Assert.Equal(2, paragraphs.Length);
            Assert.Equal("pl-PL", paragraphs[1]
                .GetFirstChild<A.EndParagraphRunProperties>()!
                .Language!.Value);
            Assert.Empty(reopened.ValidateDocument());
            Assert.Equal(bytes,
                reopened.ToBytes(PowerPointFileFormat.Ppt));
        }

        [Fact]
        public void ImportedTextLanguageAndLengthEdit_UsesIncrementalSpecialInfoRewrite() {
            byte[] sourceBytes;
            using (PowerPointPresentation source = PowerPointPresentation
                       .Create()) {
                source.AddSlide(P.SlideLayoutValues.Blank)
                    .AddTextBoxPoints("Editable language", 30, 30, 240, 60)
                    .SetLanguage("en-US");
                sourceBytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation original = LegacyPptPresentation.Load(
                sourceBytes);

            byte[] savedBytes;
            using (var input = new MemoryStream(sourceBytes,
                       writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation
                       .Load(input)) {
                PowerPointTextBox textBox = Assert.Single(
                    imported.Slides[0].TextBoxes);
                PowerPointTextRun run = Assert.Single(textBox.Paragraphs[0]
                    .Runs);
                Assert.Equal("en-US", run.Language);
                run.Language = "pl-PL";
                run.Text += "!";

                LegacyPptWritePreflightReport preflight = imported
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation saved = LegacyPptPresentation.Load(
                savedBytes);
            LegacyPptTextBody savedText = Assert.Single(Assert.Single(
                saved.Slides).Shapes).TextBody;
            Assert.Equal("Editable language!", savedText.Text);
            Assert.Collection(savedText.LanguageRuns,
                run => {
                    Assert.Equal(18, run.Length);
                    Assert.Equal("pl-PL", run.Language);
                },
                run => {
                    Assert.Equal(18, run.Start);
                    Assert.Equal(1, run.Length);
                    Assert.Equal("en-US", run.Language);
                });
            Assert.Equal(original.Package.UserEdits.Count + 1,
                saved.Package.UserEdits.Count);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));

            using var reopenedInput = new MemoryStream(savedBytes,
                writable: false);
            using PowerPointPresentation reopened = PowerPointPresentation
                .Load(reopenedInput);
            PowerPointTextBox reopenedText = Assert.Single(
                reopened.Slides[0].TextBoxes);
            Assert.Equal("pl-PL", Assert.Single(reopenedText.Paragraphs[0]
                .Runs).Language);
            P.Shape reopenedShape = Assert.IsType<P.Shape>(
                reopenedText.Element);
            Assert.Equal("en-US", reopenedShape.TextBody!
                .Elements<A.Paragraph>().Single()
                .GetFirstChild<A.EndParagraphRunProperties>()!
                .Language!.Value);
            Assert.Equal(savedBytes,
                reopened.ToBytes(PowerPointFileFormat.Ppt));
        }

        [Theory]
        [InlineData("en")]
        [InlineData("en-001")]
        public void NativeWriter_BlocksTextLanguageWithoutClassicLcid(
            string language) {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointTextBox textBox = presentation.AddSlide(
                    P.SlideLayoutValues.Blank)
                .AddTextBoxPoints("Neutral language", 30, 30, 240, 60);
            Assert.Single(textBox.Paragraphs[0].Runs).Language = language;

            LegacyPptWritePreflightReport preflight = presentation
                .AnalyzeLegacyPptWrite();

            Assert.False(preflight.CanWrite);
            LegacyPptWriteFinding finding = Assert.Single(
                preflight.Findings, item =>
                    item.Code == "PPT-WRITE-RICH-TEXT");
            Assert.Contains("no classic PowerPoint LCID mapping",
                finding.Description);
        }

        [Fact]
        public void TextSpecialInfoCodec_RejectsCustomAndTransientLcids() {
            Assert.True(LegacyPptTextSpecialInfoCodec
                .IsPersistableLanguageId(0x0409));
            Assert.False(LegacyPptTextSpecialInfoCodec
                .IsPersistableLanguageId(0x1000));
            for (int languageId = 0x2000; languageId <= 0x4C00;
                 languageId += 0x0400) {
                Assert.False(LegacyPptTextSpecialInfoCodec
                    .IsPersistableLanguageId(languageId));
            }
        }

        [Fact]
        public void ImportedUnsupportedTextSpecialInfo_AllowsUnrelatedEditAndBlocksRewrite() {
            byte[] generatedBytes;
            using (PowerPointPresentation source = PowerPointPresentation
                       .Create()) {
                PowerPointTextBox textBox = source.AddSlide(
                        P.SlideLayoutValues.Blank)
                    .AddTextBoxPoints("Preserve grammar", 30, 30, 240, 60);
                textBox.SetLanguage("en-US");
                P.Shape shape = Assert.IsType<P.Shape>(textBox.Element);
                shape.TextBody!.Descendants<A.RunProperties>().Single()
                    .SpellingError = true;
                shape.TextBody.Elements<A.Paragraph>().Single()
                    .GetFirstChild<A.EndParagraphRunProperties>()!
                    .SpellingError = true;
                generatedBytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation generated = LegacyPptPresentation.Load(
                generatedBytes);
            LegacyPptShape generatedShape = Assert.Single(Assert.Single(
                generated.Slides).Shapes);
            LegacyPptRecord shapeContainer = LegacyPptRecordReader.ReadSingle(
                generated.Package.DocumentStream,
                checked((int)generatedShape.RecordOffset),
                new LegacyPptImportOptions());
            LegacyPptRecord specialInfo = Assert.Single(shapeContainer
                .DescendantsAndSelf(), record => record.Type == 0x0FAA);
            Assert.Equal(3U, specialInfo.ReadUInt32(4));
            byte[] documentStream = (byte[])generated.Package.DocumentStream
                .Clone();
            BinaryPrimitives.WriteUInt16LittleEndian(documentStream.AsSpan(
                specialInfo.PayloadOffset + 8, 2), 0x0005);
            byte[] sourceBytes = generated.Package.RewriteCompoundStreams(
                new Dictionary<string, byte[]> {
                    ["PowerPoint Document"] = documentStream
                });

            LegacyPptPresentation original = LegacyPptPresentation.Load(
                sourceBytes);
            LegacyPptTextBody originalText = Assert.Single(Assert.Single(
                original.Slides).Shapes).TextBody;
            Assert.True(originalText.HasUnprojectedTextSpecialInfo);
            Assert.Contains(original.Diagnostics, diagnostic =>
                diagnostic.Code == "PPT-TEXT-SPECIAL-INFO-PARTIAL"
                && diagnostic.Severity
                    == LegacyPptDiagnosticSeverity.Information);

            byte[] movedBytes;
            using (var input = new MemoryStream(sourceBytes,
                       writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation
                       .Load(input)) {
                Assert.Single(imported.Slides[0].TextBoxes).Left += 15875;
                LegacyPptWritePreflightReport preflight = imported
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                movedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation moved = LegacyPptPresentation.Load(
                movedBytes);
            Assert.True(Assert.Single(Assert.Single(moved.Slides).Shapes)
                .TextBody.HasUnprojectedTextSpecialInfo);
            Assert.Equal(original.Package.UserEdits.Count + 1,
                moved.Package.UserEdits.Count);
            Assert.True(moved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));

            using var rewriteInput = new MemoryStream(sourceBytes,
                writable: false);
            using PowerPointPresentation rewrite = PowerPointPresentation
                .Load(rewriteInput);
            Assert.Single(Assert.Single(rewrite.Slides[0].TextBoxes)
                .Paragraphs[0].Runs).Language = "pl-PL";
            LegacyPptWritePreflightReport rewritePreflight = rewrite
                .AnalyzeLegacyPptWrite();
            Assert.False(rewritePreflight.CanWrite);
            Assert.Contains(rewritePreflight.Findings, finding =>
                finding.Code == "PPT-WRITE-IMPORT-LOSS");
        }

        [Fact]
        public void TextSpecialInfoReader_DecodesLcidRunsAndFlagsUnknownMetadata() {
            byte[] payload = {
                0x01, 0x00, 0x00, 0x00,
                0x02, 0x00, 0x00, 0x00,
                0x09, 0x04,
                0x02, 0x00, 0x00, 0x00,
                0x07, 0x00, 0x00, 0x00,
                0x05, 0x00,
                0x15, 0x04,
                0x0C, 0x04
            };
            var record = new LegacyPptRecord(payload, 0, 0, 0, 0x0FAA,
                0, payload.Length);

            LegacyPptTextBody result = LegacyPptTextSpecialInfoCodec.Apply(
                LegacyPptTextBody.Plain("AB"), record);

            Assert.True(result.HasTextSpecialInfoRecord);
            Assert.False(result.IsTextSpecialInfoTruncated);
            Assert.True(result.HasUnprojectedTextSpecialInfo);
            Assert.Collection(result.LanguageRuns,
                run => {
                    Assert.Equal("en-US", run.Language);
                    Assert.False(run.HasUnprojectedInformation);
                },
                run => {
                    Assert.Equal("pl-PL", run.Language);
                    Assert.Equal("fr-FR", run.AlternativeLanguage);
                    Assert.True(run.SpellingError);
                    Assert.False(run.NeedsRecheck);
                    Assert.True(run.HasUnprojectedInformation);
                });
        }

        [Fact]
        public void TextSpecialInfoReader_UsesRawTextCountPlusImplicitTerminator() {
            byte[] validPayload = {
                0x03, 0x00, 0x00, 0x00,
                0x02, 0x00, 0x00, 0x00,
                0x09, 0x04
            };
            var validRecord = new LegacyPptRecord(validPayload, 0, 0, 0,
                0x0FAA, 0, validPayload.Length);

            LegacyPptTextBody valid = LegacyPptTextSpecialInfoCodec.Apply(
                LegacyPptTextBody.Plain("A"), validRecord,
                rawCharacterCount: 2);

            Assert.False(valid.IsTextSpecialInfoTruncated);
            Assert.True(valid.HasUnprojectedTextSpecialInfo);
            Assert.Equal(3, Assert.Single(valid.LanguageRuns).Length);

            byte[] shortPayload = {
                0x02, 0x00, 0x00, 0x00,
                0x02, 0x00, 0x00, 0x00,
                0x09, 0x04
            };
            var shortRecord = new LegacyPptRecord(shortPayload, 0, 0, 0,
                0x0FAA, 0, shortPayload.Length);

            LegacyPptTextBody truncated = LegacyPptTextSpecialInfoCodec.Apply(
                LegacyPptTextBody.Plain("A"), shortRecord,
                rawCharacterCount: 2);

            Assert.True(truncated.IsTextSpecialInfoTruncated);
        }

        [Fact]
        public void TextSpecialInfoReader_LossGatesMixedZeroAndTransientLcids() {
            byte[] mixedPayload = {
                0x01, 0x00, 0x00, 0x00,
                0x02, 0x00, 0x00, 0x00,
                0x00, 0x00,
                0x01, 0x00, 0x00, 0x00,
                0x02, 0x00, 0x00, 0x00,
                0x09, 0x04
            };
            var mixedRecord = new LegacyPptRecord(mixedPayload, 0, 0, 0,
                0x0FAA, 0, mixedPayload.Length);

            LegacyPptTextBody mixed = LegacyPptTextSpecialInfoCodec.Apply(
                LegacyPptTextBody.Plain("A"), mixedRecord);

            Assert.False(mixed.IsTextSpecialInfoTruncated);
            Assert.True(mixed.HasUnprojectedTextSpecialInfo);
            Assert.Equal((ushort)0, mixed.LanguageRuns[0].LanguageId);
            Assert.Equal("en-US", mixed.LanguageRuns[1].Language);

            byte[] transientPayload = {
                0x01, 0x00, 0x00, 0x00,
                0x02, 0x00, 0x00, 0x00,
                0x00, 0x20
            };
            var transientRecord = new LegacyPptRecord(transientPayload, 0,
                0, 0, 0x0FAA, 0, transientPayload.Length);

            LegacyPptTextBody transient = LegacyPptTextSpecialInfoCodec.Apply(
                LegacyPptTextBody.Plain(string.Empty), transientRecord);

            Assert.False(transient.IsTextSpecialInfoTruncated);
            Assert.True(transient.HasUnprojectedTextSpecialInfo);
            LegacyPptTextLanguageRun transientRun = Assert.Single(
                transient.LanguageRuns);
            Assert.Equal((ushort)0x2000, transientRun.LanguageId);
            Assert.Null(transientRun.Language);
            Assert.True(transientRun.HasUnprojectedInformation);
        }
    }
}
