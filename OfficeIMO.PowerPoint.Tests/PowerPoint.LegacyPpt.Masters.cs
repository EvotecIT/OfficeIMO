using System.Buffers.Binary;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing.Binary;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using OfficeIMO.Tests.Pdf;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class PowerPointLegacyPptMasterTests {
        private static string FixturePath => Path.Combine(AppContext.BaseDirectory,
            "Documents", "LegacyPptCorpus", "BasicPowerPoint.ppt");

        [Fact]
        public void DocumentAtomReader_DecodesCompleteDocumentAndMasterSettings() {
            var payload = new byte[40];
            BinaryPrimitives.WriteInt32LittleEndian(payload.AsSpan(0, 4), 7200);
            BinaryPrimitives.WriteInt32LittleEndian(payload.AsSpan(4, 4), 5400);
            BinaryPrimitives.WriteInt32LittleEndian(payload.AsSpan(8, 4), 5400);
            BinaryPrimitives.WriteInt32LittleEndian(payload.AsSpan(12, 4), 7200);
            BinaryPrimitives.WriteInt32LittleEndian(payload.AsSpan(16, 4), 3);
            BinaryPrimitives.WriteInt32LittleEndian(payload.AsSpan(20, 4), 2);
            BinaryPrimitives.WriteUInt32LittleEndian(payload.AsSpan(24, 4), 11);
            BinaryPrimitives.WriteUInt32LittleEndian(payload.AsSpan(28, 4), 12);
            BinaryPrimitives.WriteUInt16LittleEndian(payload.AsSpan(32, 2), 7);
            BinaryPrimitives.WriteUInt16LittleEndian(payload.AsSpan(34, 2),
                (ushort)LegacyPptSlideSizeType.A4Paper);
            payload[36] = 1;
            payload[37] = 1;
            payload[38] = 1;
            payload[39] = 1;
            var record = new LegacyPptRecord(payload, 0, 1, 0, 0x03E9,
                0, payload.Length);

            LegacyPptDocumentSettings settings = Assert.IsType<LegacyPptDocumentSettings>(
                LegacyPptDocumentAtomReader.Read(record));

            Assert.Equal(7200, settings.SlideWidth);
            Assert.Equal(5400, settings.SlideHeight);
            Assert.Equal(5400, settings.NotesWidth);
            Assert.Equal(7200, settings.NotesHeight);
            Assert.Equal(3, settings.ServerZoomNumerator);
            Assert.Equal(2, settings.ServerZoomDenominator);
            Assert.Equal(11U, settings.NotesMasterPersistId);
            Assert.Equal(12U, settings.HandoutMasterPersistId);
            Assert.Equal(7, settings.FirstSlideNumber);
            Assert.Equal(LegacyPptSlideSizeType.A4Paper, settings.SlideSizeType);
            Assert.True(settings.SaveWithFonts);
            Assert.True(settings.OmitTitlePlaceholders);
            Assert.True(settings.RightToLeft);
            Assert.True(settings.ShowComments);
            Assert.Null(LegacyPptDocumentAtomReader.Read(new LegacyPptRecord(
                new byte[39], 0, 1, 0, 0x03E9, 0, 39)));
        }

        [Fact]
        public void BinaryImport_ProjectsDocumentSettingsAndNotesMasterTopology() {
            LegacyPptPresentation legacy = LegacyPptPresentation.Load(FixturePath);
            LegacyPptDocumentSettings settings = Assert.IsType<LegacyPptDocumentSettings>(
                legacy.DocumentSettings);
            Assert.Equal(settings.NotesMasterPersistId != 0, legacy.NotesMaster != null);
            if (legacy.NotesMaster != null) {
                Assert.Equal(LegacyPptSpecialMasterKind.Notes, legacy.NotesMaster.Kind);
                Assert.Equal(settings.NotesMasterPersistId, legacy.NotesMaster.PersistId);
                Assert.NotNull(legacy.NotesMaster.ColorScheme);
                Assert.NotEmpty(legacy.NotesMaster.Shapes);
            }

            using PowerPointPresentation projected = PowerPointPresentation.Load(FixturePath);
            P.Presentation root = projected.OpenXmlDocument.PresentationPart!.Presentation;
            Assert.Equal(settings.FirstSlideNumber, root.FirstSlideNum?.Value);
            Assert.Equal(checked((int)Math.Round(100000D
                    * settings.ServerZoomNumerator / settings.ServerZoomDenominator,
                MidpointRounding.AwayFromZero)), root.ServerZoom?.Value);
            Assert.Equal(!settings.OmitTitlePlaceholders,
                root.ShowSpecialPlaceholderOnTitleSlide?.Value);
            Assert.Equal(settings.RightToLeft, root.RightToLeft?.Value);
            Assert.Equal(settings.SaveWithFonts, root.EmbedTrueTypeFonts?.Value);
            Assert.Equal(settings.ShowComments, projected.OpenXmlDocument.PresentationPart!
                .ViewPropertiesPart!.ViewProperties!.ShowComments?.Value);
            Assert.Equal(ToEmus(settings.NotesWidth), root.NotesSize!.Cx!.Value);
            Assert.Equal(ToEmus(settings.NotesHeight), root.NotesSize.Cy!.Value);

            if (legacy.NotesMaster != null) {
                NotesMasterPart notesPart = projected.OpenXmlDocument.PresentationPart!
                    .NotesMasterPart!;
                Assert.Equal("Binary Notes Master",
                    notesPart.NotesMaster!.CommonSlideData!.Name!.Value);
                Assert.NotNull(notesPart.ThemePart?.Theme?.ThemeElements?.ColorScheme);
                Assert.NotEmpty(notesPart.NotesMaster.CommonSlideData.ShapeTree!
                    .Descendants<P.PlaceholderShape>());
            }
            Assert.Empty(projected.ValidateDocument());

            LegacyPptImportReport report = legacy.CreateImportReport();
            Assert.Equal(legacy.NotesMaster == null ? 0 : 1, report.SpecialMasterCount);
            Assert.Equal(legacy.NotesMaster?.Shapes.Count ?? 0,
                report.SpecialMasterShapeCount);
        }

        [Fact]
        public void NativeWriter_RoundTripsDocumentPageAndDisplaySettings() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            presentation.SlideSize.SetSizeEmus(9906000, 6858000, P.SlideSizeValues.A4);
            P.Presentation root = presentation.OpenXmlDocument.PresentationPart!.Presentation;
            root.NotesSize = new P.NotesSize { Cx = 6858000, Cy = 9906000 };
            root.FirstSlideNum = 7;
            root.ShowSpecialPlaceholderOnTitleSlide = false;
            root.RightToLeft = true;
            root.EmbedTrueTypeFonts = true;
            root.ServerZoom = 75000;
            presentation.OpenXmlDocument.PresentationPart!.ViewPropertiesPart!
                .ViewProperties!.ShowComments = true;
            presentation.AddSlide();

            LegacyPptDocumentSettings settings = Assert.IsType<LegacyPptDocumentSettings>(
                LegacyPptPresentation.Load(presentation.ToBytes(PowerPointFileFormat.Ppt))
                    .DocumentSettings);

            Assert.Equal(LegacyPptSlideSizeType.A4Paper, settings.SlideSizeType);
            Assert.Equal(7, settings.FirstSlideNumber);
            Assert.True(settings.OmitTitlePlaceholders);
            Assert.True(settings.RightToLeft);
            Assert.True(settings.SaveWithFonts);
            Assert.True(settings.ShowComments);
            Assert.Equal(3, settings.ServerZoomNumerator);
            Assert.Equal(4, settings.ServerZoomDenominator);
            Assert.Equal(6858000, ToEmus(settings.NotesWidth));
            Assert.Equal(9906000, ToEmus(settings.NotesHeight));
        }

        [Fact]
        public void NativeWriter_WritesReferencedMasterTopologyAndClassicColorSchemes() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            presentation.SetThemeColors(new Dictionary<PowerPointThemeColor, string> {
                [PowerPointThemeColor.Light1] = "F1F2F3",
                [PowerPointThemeColor.Dark1] = "111213",
                [PowerPointThemeColor.Accent4] = "414243",
                [PowerPointThemeColor.Dark2] = "212223",
                [PowerPointThemeColor.Light2] = "E1E2E3",
                [PowerPointThemeColor.Accent1] = "A1A2A3",
                [PowerPointThemeColor.Accent2] = "B1B2B3",
                [PowerPointThemeColor.Accent3] = "C1C2C3"
            });
            presentation.AddSlide();

            byte[] bytes = presentation.ToBytes(PowerPointFileFormat.Ppt);
            LegacyPptPresentation binary = LegacyPptPresentation.Load(bytes);
            LegacyPptMaster master = Assert.Single(binary.Masters);
            LegacyPptColorScheme scheme = Assert.IsType<LegacyPptColorScheme>(
                master.ColorScheme);

            Assert.Equal(master.MasterId, Assert.Single(binary.Slides).MasterId);
            Assert.Equal("F1F2F3", scheme.Background);
            Assert.Equal("111213", scheme.Text);
            Assert.Equal("414243", scheme.Shadow);
            Assert.Equal("212223", scheme.TitleText);
            Assert.Equal("E1E2E3", scheme.Fill);
            Assert.Equal("A1A2A3", scheme.Accent1);
            Assert.Equal("B1B2B3", scheme.Accent2);
            Assert.Equal("C1C2C3", scheme.Accent3);

            using var stream = new MemoryStream(bytes);
            using PowerPointPresentation reopened = PowerPointPresentation.Load(stream);
            Assert.Equal("F1F2F3",
                reopened.GetThemeColor(PowerPointThemeColor.Light1));
            Assert.Equal("111213",
                reopened.GetThemeColor(PowerPointThemeColor.Dark1));
            Assert.Equal("414243",
                reopened.GetThemeColor(PowerPointThemeColor.Accent4));
            Assert.Empty(reopened.ValidateDocument());
        }

        [Fact]
        public void NativeWriter_PreservesDistinctOpenXmlSlideMasters() {
            using PowerPointPresentation target = PowerPointPresentation.Create();
            target.SetThemeColor(PowerPointThemeColor.Accent1, "102030");
            target.AddSlide();

            using PowerPointPresentation source = PowerPointPresentation.Create();
            source.SetThemeColor(PowerPointThemeColor.Accent1, "A0B0C0");
            source.OpenXmlDocument.PresentationPart!.SlideMasterParts.First()
                .ThemePart!.Theme!.Save();
            SlideLayoutPart sourceLayout = source.OpenXmlDocument.PresentationPart!
                .SlideMasterParts.First().SlideLayoutParts.First();
            sourceLayout.SlideLayout!.CommonSlideData!.Name = "Imported unique layout";
            sourceLayout.SlideLayout.Save();
            source.AddSlide();
            target.ImportSlide(source, 0);

            Assert.Equal(2, target.OpenXmlDocument.PresentationPart!
                .SlideMasterParts.Count());
            byte[] bytes = target.ToBytes(PowerPointFileFormat.Ppt);
            LegacyPptPresentation binary = LegacyPptPresentation.Load(bytes);

            Assert.Equal(2, binary.Masters.Count);
            Assert.Equal(2, binary.Slides.Select(slide => slide.MasterId)
                .Distinct().Count());
            Assert.Equal("102030", binary.Masters[0].ColorScheme!.Accent1);
            Assert.Equal("A0B0C0", binary.Masters[1].ColorScheme!.Accent1);

            using var stream = new MemoryStream(bytes);
            using PowerPointPresentation reopened = PowerPointPresentation.Load(stream);
            Assert.Equal(2, reopened.OpenXmlDocument.PresentationPart!
                .SlideMasterParts.Count());
            Assert.Equal("102030",
                reopened.GetThemeColor(PowerPointThemeColor.Accent1, 0));
            Assert.Equal("A0B0C0",
                reopened.GetThemeColor(PowerPointThemeColor.Accent1, 1));
            Assert.Empty(reopened.ValidateDocument());
        }

        [Fact]
        public void NativeWriter_ExpandsBeyondEmbeddedSlideMasterScaffold() {
            const int masterCount = 12;
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            presentation.SetThemeColor(PowerPointThemeColor.Accent1, "010101");
            presentation.AddSlide();

            for (int index = 1; index < masterCount; index++) {
                using PowerPointPresentation source = PowerPointPresentation.Create();
                source.SetThemeColor(PowerPointThemeColor.Accent1,
                    $"{index + 1:X2}{index + 1:X2}{index + 1:X2}");
                SlideLayoutPart layout = source.OpenXmlDocument.PresentationPart!
                    .SlideMasterParts.First().SlideLayoutParts.First();
                layout.SlideLayout!.CommonSlideData!.Name =
                    $"Imported master {index + 1}";
                layout.SlideLayout.Save();
                source.AddSlide();
                presentation.ImportSlide(source, 0);
            }
            presentation.Slides[0].Notes.Text = "Expanded topology note";
            CreateHandoutMaster(presentation);

            Assert.Equal(masterCount, presentation.OpenXmlDocument
                .PresentationPart!.SlideMasterParts.Count());
            LegacyPptWritePreflightReport preflight =
                presentation.AnalyzeLegacyPptWrite();
            Assert.True(preflight.CanWrite,
                string.Join(Environment.NewLine, preflight.Findings));

            byte[] bytes = presentation.ToBytes(PowerPointFileFormat.Ppt);
            LegacyPptPresentation binary = LegacyPptPresentation.Load(bytes);

            Assert.Equal(masterCount, binary.Masters.Count);
            Assert.Equal(masterCount, binary.Slides.Count);
            Assert.Equal(masterCount, binary.Slides.Select(slide => slide.MasterId)
                .Distinct().Count());
            Assert.Equal(14U, binary.DocumentSettings!.NotesMasterPersistId);
            Assert.Equal(28U, binary.DocumentSettings.HandoutMasterPersistId);
            Assert.Equal(Enumerable.Range(15, masterCount).Select(value =>
                    unchecked((uint)value)),
                binary.Slides.Select(slide => slide.PersistId));
            Assert.Equal(27U, Assert.IsType<LegacyPptNotesPage>(
                binary.Slides[0].NotesPage).PersistId);
            Assert.Equal(28U, Assert.IsType<LegacyPptSpecialMaster>(
                binary.HandoutMaster).PersistId);

            using var stream = new MemoryStream(bytes);
            using PowerPointPresentation reopened = PowerPointPresentation.Load(stream);
            Assert.Equal(masterCount, reopened.OpenXmlDocument.PresentationPart!
                .SlideMasterParts.Count());
            Assert.Equal(masterCount, reopened.Slides.Count);
            Assert.Equal("Expanded topology note", reopened.Slides[0].Notes.Text);
            Assert.NotNull(reopened.OpenXmlDocument.PresentationPart!
                .HandoutMasterPart);
            Assert.Empty(reopened.ValidateDocument());
        }

        [Fact]
        public void NativeWriter_WritesMainMasterShapesAndPlaceholderKinds() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            SlideMasterPart masterPart = presentation.OpenXmlDocument
                .PresentationPart!.SlideMasterParts.First();
            P.ShapeTree tree = masterPart.SlideMaster!.CommonSlideData!.ShapeTree!;
            var titleBounds = new PowerPointLayoutBox(650000, 400000,
                7600000, 900000);
            tree.Append(
                new P.Shape(
                    new P.NonVisualShapeProperties(
                        new P.NonVisualDrawingProperties {
                            Id = 2U,
                            Name = "Master title"
                        },
                        new P.NonVisualShapeDrawingProperties(),
                        new P.ApplicationNonVisualDrawingProperties(
                            new P.PlaceholderShape {
                                Type = P.PlaceholderValues.Title,
                                Index = 0U
                            })),
                    new P.ShapeProperties(
                        new A.Transform2D(
                            new A.Offset {
                                X = titleBounds.Left,
                                Y = titleBounds.Top
                            },
                            new A.Extents {
                                Cx = titleBounds.Width,
                                Cy = titleBounds.Height
                            }),
                        new A.PresetGeometry(new A.AdjustValueList()) {
                            Preset = A.ShapeTypeValues.Rectangle
                        }),
                    new P.TextBody(
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        new A.Paragraph(
                            new A.Run(new A.Text("Master title")),
                            new A.EndParagraphRunProperties()))),
                new P.Shape(
                    new P.NonVisualShapeProperties(
                        new P.NonVisualDrawingProperties {
                            Id = 3U,
                            Name = "Master decoration"
                        },
                        new P.NonVisualShapeDrawingProperties(),
                        new P.ApplicationNonVisualDrawingProperties()),
                    new P.ShapeProperties(
                        new A.Transform2D(
                            new A.Offset { X = 300000, Y = 6000000 },
                            new A.Extents { Cx = 8500000, Cy = 180000 }),
                        new A.PresetGeometry(new A.AdjustValueList()) {
                            Preset = A.ShapeTypeValues.Rectangle
                        })));
            presentation.AddSlide(P.SlideLayoutValues.Blank);

            LegacyPptWritePreflightReport preflight = presentation
                .AnalyzeLegacyPptWrite();
            Assert.True(preflight.CanWrite,
                string.Join(Environment.NewLine, preflight.Findings));
            byte[] bytes = presentation.ToBytes(PowerPointFileFormat.Ppt);
            LegacyPptMaster master = Assert.Single(
                LegacyPptPresentation.Load(bytes).Masters);

            Assert.Equal(2, master.Shapes.Count);
            LegacyPptShape title = Assert.Single(master.Shapes, shape =>
                shape.Placeholder?.Kind == LegacyPptPlaceholderKind.MasterTitle);
            Assert.Equal("Master title", title.Text);
            Assert.Equal(LegacyPptTextType.Title, title.TextBody.TextType);
            Assert.Equal(ToEmus(ToMasterUnits(titleBounds.Left)),
                ToEmus(title.Bounds.Left));
            Assert.Equal(ToEmus(ToMasterUnits(titleBounds.Top)),
                ToEmus(title.Bounds.Top));
            LegacyPptShape decoration = Assert.Single(master.Shapes,
                shape => shape.Placeholder == null);
            Assert.Equal(LegacyPptShapeKind.Rectangle, decoration.Kind);
            Assert.Equal(title.ShapeId >> 10, decoration.ShapeId >> 10);

            using var stream = new MemoryStream(bytes);
            using PowerPointPresentation reopened = PowerPointPresentation.Load(
                stream);
            P.Shape[] masterShapes = reopened.OpenXmlDocument.PresentationPart!
                .SlideMasterParts.Single().SlideMaster!.CommonSlideData!.ShapeTree!
                .Elements<P.Shape>().ToArray();
            Assert.Equal(2, masterShapes.Length);
            Assert.Contains(masterShapes, shape => shape.NonVisualShapeProperties?
                .ApplicationNonVisualDrawingProperties?.PlaceholderShape?
                .Type?.Value == P.PlaceholderValues.Title);
            PowerPointShape inheritedDecoration = Assert.Single(reopened.Slides[0]
                .GetInheritedShapesForExport());
            Assert.Equal("Master decoration", inheritedDecoration.Name);
            Assert.Empty(reopened.ValidateDocument());
        }

        [Fact]
        public void BinaryImport_ProjectsMasterFieldsAndLanguageAsRichText() {
            byte[] bytes;
            using (PowerPointPresentation source =
                   PowerPointPresentation.Create()) {
                SlideMasterPart masterPart = source.OpenXmlDocument
                    .PresentationPart!.SlideMasterParts.Single();
                P.ShapeTree tree = masterPart.SlideMaster!
                    .CommonSlideData!.ShapeTree!;
                tree.Append(
                    new P.Shape(
                        new P.NonVisualShapeProperties(
                            new P.NonVisualDrawingProperties {
                                Id = 800U,
                                Name = "Master dynamic field"
                            },
                            new P.NonVisualShapeDrawingProperties(),
                            new P.ApplicationNonVisualDrawingProperties()),
                        new P.ShapeProperties(
                            new A.Transform2D(
                                new A.Offset { X = 100000, Y = 100000 },
                                new A.Extents { Cx = 800000, Cy = 300000 }),
                            new A.PresetGeometry(new A.AdjustValueList()) {
                                Preset = A.ShapeTypeValues.Rectangle
                            }),
                        new P.TextBody(new A.BodyProperties(),
                            new A.ListStyle(), new A.Paragraph(
                                new A.Field(new A.Text("1")) {
                                    Id = "{00000000-0000-0000-0000-000000000801}",
                                    Type = "slidenum"
                                }))),
                    new P.Shape(
                        new P.NonVisualShapeProperties(
                            new P.NonVisualDrawingProperties {
                                Id = 801U,
                                Name = "Master language"
                            },
                            new P.NonVisualShapeDrawingProperties(),
                            new P.ApplicationNonVisualDrawingProperties()),
                        new P.ShapeProperties(
                            new A.Transform2D(
                                new A.Offset { X = 100000, Y = 500000 },
                                new A.Extents { Cx = 1200000, Cy = 300000 }),
                            new A.PresetGeometry(new A.AdjustValueList()) {
                                Preset = A.ShapeTypeValues.Rectangle
                            }),
                        new P.TextBody(new A.BodyProperties(),
                            new A.ListStyle(), new A.Paragraph(
                                new A.Run(new A.RunProperties {
                                    Language = "pl-PL"
                                }, new A.Text("Język"))))));
                source.AddSlide(P.SlideLayoutValues.Blank);
                LegacyPptWritePreflightReport preflight = source
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                bytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptMaster binaryMaster = Assert.Single(
                LegacyPptPresentation.Load(bytes).Masters);
            Assert.Contains(binaryMaster.Shapes, shape =>
                shape.TextBody.Fields.Count == 1);
            Assert.Contains(binaryMaster.Shapes, shape =>
                shape.TextBody.HasLanguageInformation);

            using var input = new MemoryStream(bytes, writable: false);
            using PowerPointPresentation projected =
                PowerPointPresentation.Load(input);
            SlideMasterPart projectedMaster = projected.OpenXmlDocument
                .PresentationPart!.SlideMasterParts.Single();
            Assert.Single(projectedMaster.SlideMaster!
                .Descendants<A.Field>(), field =>
                field.Type?.Value == "slidenum");
            Assert.Contains(projectedMaster.SlideMaster!
                .Descendants<A.RunProperties>(), properties =>
                properties.Language?.Value == "pl-PL");
            Assert.Empty(projected.ValidateDocument());
        }

        [Fact]
        public void NativeWriter_RoundTripsMainMasterBaseTextStylesAndFonts() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            SlideMasterPart masterPart = presentation.OpenXmlDocument
                .PresentationPart!.SlideMasterParts.First();
            var titleLevel = new A.Level1ParagraphProperties {
                Alignment = A.TextAlignmentTypeValues.Center,
                LeftMargin = 317500,
                Indent = -158750,
                DefaultTabSize = 635000,
                FontAlignment = A.TextFontAlignmentValues.Center,
                RightToLeft = true,
                EastAsianLineBreak = true,
                LatinLineBreak = false,
                Height = true
            };
            titleLevel.Append(
                new A.LineSpacing(new A.SpacingPercent { Val = 125000 }),
                new A.SpaceBefore(new A.SpacingPoints { Val = 1250 }),
                new A.SpaceAfter(new A.SpacingPoints { Val = 13 }),
                new A.BulletColor(
                    new A.RgbColorModelHex { Val = "123456" }),
                new A.BulletSizePercentage { Val = 120000 },
                new A.BulletFont { Typeface = "OfficeIMO Bullet" },
                new A.CharacterBullet { Char = "•" },
                new A.TabStopList(new A.TabStop {
                    Position = 476250,
                    Alignment = A.TextTabAlignmentValues.Decimal
                }),
                new A.DefaultRunProperties(
                    new A.SolidFill(
                        new A.RgbColorModelHex { Val = "654321" }),
                    new A.LatinFont { Typeface = "OfficeIMO Latin" },
                    new A.EastAsianFont { Typeface = "OfficeIMO East" },
                    new A.SymbolFont { Typeface = "OfficeIMO Symbol" }) {
                    Bold = true,
                    Italic = true,
                    Underline = A.TextUnderlineValues.Single,
                    Kumimoji = true,
                    FontSize = 3200,
                    Baseline = 10000
                });
            masterPart.SlideMaster!.TextStyles = new P.TextStyles(
                new P.TitleStyle(
                    titleLevel,
                    new A.Level2ParagraphProperties(
                        new A.DefaultRunProperties(
                            new A.SolidFill(new A.SchemeColor {
                                Val = A.SchemeColorValues.Accent1
                            }))),
                    new A.Level3ParagraphProperties(),
                    new A.Level4ParagraphProperties(),
                    new A.Level5ParagraphProperties()),
                new P.BodyStyle(),
                new P.OtherStyle());
            presentation.AddSlide(P.SlideLayoutValues.Blank);

            LegacyPptWritePreflightReport preflight = presentation
                .AnalyzeLegacyPptWrite();
            Assert.True(preflight.CanWrite,
                string.Join(Environment.NewLine, preflight.Findings));
            byte[] bytes = presentation.ToBytes(PowerPointFileFormat.Ppt);
            LegacyPptPresentation binary = LegacyPptPresentation.Load(bytes);
            LegacyPptMaster binaryMaster = Assert.Single(binary.Masters);
            Assert.Equal(3, binaryMaster.TextMasterStyles.Count);
            LegacyPptTextMasterStyle style = Assert.Single(
                binaryMaster.TextMasterStyles,
                candidate => candidate.TextType == LegacyPptTextType.Title);
            Assert.Equal(5, style.Levels.Count);
            LegacyPptTextMasterStyleLevel level = style.Levels[0];
            LegacyPptParagraphRun paragraph = level.ParagraphProperties;
            LegacyPptCharacterRun character = level.CharacterProperties;

            Assert.False(style.IsTruncated);
            Assert.False(style.HasUnprojectedFormatting,
                $"paragraph={paragraph.HasUnprojectedFormatting}; "
                + $"character={character.HasUnprojectedFormatting}; "
                + $"fonts={string.Join(",", binary.Fonts.Select(font => $"{font.Index}:{font.Typeface}"))}");
            Assert.True(paragraph.HasBullet);
            Assert.True(paragraph.BulletHasFont);
            Assert.True(paragraph.BulletHasColor);
            Assert.True(paragraph.BulletHasSize);
            Assert.Equal('•', paragraph.BulletCharacter);
            Assert.Equal("OfficeIMO Bullet", paragraph.BulletTypeface);
            Assert.Equal((short)120, paragraph.BulletSize);
            Assert.Equal("123456", paragraph.BulletColor);
            Assert.Equal(LegacyPptTextAlignment.Center, paragraph.Alignment);
            Assert.Equal((short)125, paragraph.LineSpacing);
            Assert.Equal((short)-100, paragraph.SpaceBefore);
            Assert.Equal((short)-1, paragraph.SpaceAfter);
            Assert.Equal((short)200, paragraph.LeftMargin);
            Assert.Equal((short)-100, paragraph.Indent);
            Assert.Equal((short)400, paragraph.DefaultTabSize);
            Assert.Equal(LegacyPptFontAlignment.Center,
                paragraph.FontAlignment);
            Assert.True(paragraph.CharacterWrap);
            Assert.True(paragraph.WordWrap);
            Assert.True(paragraph.Overflow);
            Assert.Equal(LegacyPptTextDirection.RightToLeft,
                paragraph.TextDirection);
            LegacyPptTabStop tab = Assert.Single(paragraph.TabStops);
            Assert.Equal((short)300, tab.Position);
            Assert.Equal(LegacyPptTabAlignment.Decimal, tab.Alignment);
            Assert.True(character.Bold);
            Assert.True(character.Italic);
            Assert.True(character.Underline);
            Assert.True(character.Kumi);
            Assert.Equal((short)32, character.FontSizePoints);
            Assert.Equal((short)10, character.BaselinePositionPercent);
            Assert.Equal("654321", character.Color);
            Assert.Equal("OfficeIMO Latin", character.Typeface);
            Assert.Equal("OfficeIMO East", character.OldEastAsianTypeface);
            Assert.Equal("OfficeIMO Symbol", character.SymbolTypeface);
            Assert.Equal((byte)5,
                style.Levels[1].CharacterProperties.ColorSchemeIndex);
            Assert.All(new[] { "OfficeIMO Bullet", "OfficeIMO Latin",
                    "OfficeIMO East", "OfficeIMO Symbol" }, typeface =>
                Assert.Contains(binary.Fonts, font => font.Typeface == typeface));

            using var stream = new MemoryStream(bytes);
            using PowerPointPresentation reopened = PowerPointPresentation.Load(
                stream);
            A.Level1ParagraphProperties projected = Assert.IsType<
                A.Level1ParagraphProperties>(reopened.OpenXmlDocument
                .PresentationPart!.SlideMasterParts.Single().SlideMaster!
                .TextStyles!.TitleStyle!.FirstChild);
            Assert.Equal(317500, projected.LeftMargin!.Value);
            Assert.Equal(-158750, projected.Indent!.Value);
            Assert.Equal(635000, projected.DefaultTabSize!.Value);
            Assert.Equal(A.TextAlignmentTypeValues.Center,
                projected.Alignment!.Value);
            Assert.Equal("OfficeIMO Bullet",
                projected.GetFirstChild<A.BulletFont>()!.Typeface!.Value);
            Assert.Equal("•",
                projected.GetFirstChild<A.CharacterBullet>()!.Char!.Value);
            A.DefaultRunProperties projectedRun = projected
                .GetFirstChild<A.DefaultRunProperties>()!;
            Assert.Equal("OfficeIMO Latin",
                projectedRun.GetFirstChild<A.LatinFont>()!.Typeface!.Value);
            Assert.Equal("OfficeIMO East", projectedRun
                .GetFirstChild<A.EastAsianFont>()!.Typeface!.Value);
            Assert.Equal("OfficeIMO Symbol", projectedRun
                .GetFirstChild<A.SymbolFont>()!.Typeface!.Value);
            A.Level2ParagraphProperties projectedSecond = Assert.IsType<
                A.Level2ParagraphProperties>(reopened.OpenXmlDocument
                .PresentationPart!.SlideMasterParts.Single().SlideMaster!
                .TextStyles!.TitleStyle!.ChildElements[1]);
            Assert.Equal(A.SchemeColorValues.Accent1, projectedSecond
                .GetFirstChild<A.DefaultRunProperties>()!
                .GetFirstChild<A.SolidFill>()!.SchemeColor!.Val!.Value);
            Assert.Empty(reopened.ValidateDocument());
        }

        [Fact]
        public void NativeWriter_RejectsUnrepresentableMainMasterTextStyle() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            SlideMasterPart masterPart = presentation.OpenXmlDocument
                .PresentationPart!.SlideMasterParts.First();
            var titleLevel = new A.Level1ParagraphProperties(
                new A.DefaultRunProperties { FontSize = 3250 });
            masterPart.SlideMaster!.TextStyles = new P.TextStyles(
                new P.TitleStyle(titleLevel),
                new P.BodyStyle(),
                new P.OtherStyle());
            presentation.AddSlide(P.SlideLayoutValues.Blank);

            LegacyPptWritePreflightReport preflight = presentation
                .AnalyzeLegacyPptWrite();

            Assert.False(preflight.CanWrite);
            Assert.Contains(preflight.Findings, finding =>
                finding.Code == "PPT-WRITE-MASTER-TEXT-STYLE");
        }

        [Fact]
        public void NativeWriter_RoundTripsSolidGradientAndNoFillSlideBackgrounds() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            PowerPointSlide solid = presentation.AddSlide(P.SlideLayoutValues.Blank);
            solid.BackgroundColor = "123456";
            PowerPointSlide gradient = presentation.AddSlide(P.SlideLayoutValues.Blank);
            gradient.SetBackgroundGradient("112233", "AABBCC", 45D);
            PowerPointSlide noFill = presentation.AddSlide(P.SlideLayoutValues.Blank);
            noFill.SlidePart.Slide!.CommonSlideData!.Background = new P.Background(
                new P.BackgroundProperties(new A.NoFill()));

            Assert.True(presentation.AnalyzeLegacyPptWrite().CanWrite);
            byte[] bytes = presentation.ToBytes(PowerPointFileFormat.Ppt);
            LegacyPptPresentation binary = LegacyPptPresentation.Load(bytes);

            Assert.Collection(binary.Slides,
                slide => {
                    Assert.False(slide.FollowsMasterBackground);
                    LegacyPptBackground background = Assert.IsType<LegacyPptBackground>(
                        slide.Background);
                    Assert.Equal(LegacyPptBackgroundKind.Solid, background.Kind);
                    Assert.Equal("123456", background.ForegroundColor);
                },
                slide => {
                    Assert.False(slide.FollowsMasterBackground);
                    LegacyPptBackground background = Assert.IsType<LegacyPptBackground>(
                        slide.Background);
                    Assert.Equal(LegacyPptBackgroundKind.LinearGradient,
                        background.Kind);
                    Assert.Equal("112233", background.ForegroundColor);
                    Assert.Equal("AABBCC", background.BackgroundColor);
                    Assert.Equal(225D, background.AngleDegrees);
                    Assert.Equal(2, background.GradientStops.Count);
                },
                slide => {
                    Assert.False(slide.FollowsMasterBackground);
                    Assert.Equal(LegacyPptBackgroundKind.None,
                        Assert.IsType<LegacyPptBackground>(slide.Background).Kind);
                });

            using var stream = new MemoryStream(bytes);
            using PowerPointPresentation reopened = PowerPointPresentation.Load(stream);
            PowerPointSlideBackground solidBackground = reopened.Slides[0].GetBackground();
            Assert.Equal(PowerPointSlideBackgroundKind.SolidColor,
                solidBackground.Kind);
            Assert.Equal("123456", solidBackground.Color);
            PowerPointSlideBackground gradientBackground = reopened.Slides[1].GetBackground();
            Assert.Equal(PowerPointSlideBackgroundKind.LinearGradient,
                gradientBackground.Kind);
            Assert.Equal("112233", gradientBackground.GradientStartColor);
            Assert.Equal("AABBCC", gradientBackground.GradientEndColor);
            Assert.Equal(45D, gradientBackground.GradientAngleDegrees);
            Assert.NotNull(reopened.Slides[2].SlidePart.Slide!.CommonSlideData!
                .Background!.BackgroundProperties!.GetFirstChild<A.NoFill>());
            Assert.Empty(reopened.ValidateDocument());
        }

        [Fact]
        public void NativeWriter_WritesMasterAndMaterializedLayoutBackgrounds() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            SlideMasterPart masterPart = presentation.OpenXmlDocument.PresentationPart!
                .SlideMasterParts.First();
            masterPart.SlideMaster!.CommonSlideData!.Background = new P.Background(
                new P.BackgroundProperties(
                    new A.SolidFill(new A.RgbColorModelHex { Val = "0A0B0C" })));
            PowerPointSlide inherited = presentation.AddSlide(P.SlideLayoutValues.Title);
            int blankIndex = presentation.GetLayoutIndex(P.SlideLayoutValues.Blank);
            SlideLayoutPart blankLayout = masterPart.SlideLayoutParts
                .ElementAt(blankIndex);
            blankLayout.SlideLayout!.CommonSlideData!.Background = new P.Background(
                new P.BackgroundProperties(
                    new A.SolidFill(new A.RgbColorModelHex { Val = "8899AA" })));
            PowerPointSlide materialized = presentation.AddSlide(0, blankIndex);

            byte[] bytes = presentation.ToBytes(PowerPointFileFormat.Ppt);
            LegacyPptPresentation binary = LegacyPptPresentation.Load(bytes);
            LegacyPptMaster master = Assert.Single(binary.Masters);
            Assert.Equal("0A0B0C",
                Assert.IsType<LegacyPptBackground>(master.Background).ForegroundColor);
            Assert.True(binary.Slides[0].FollowsMasterBackground);
            Assert.False(binary.Slides[1].FollowsMasterBackground);
            Assert.Equal("8899AA", Assert.IsType<LegacyPptBackground>(
                binary.Slides[1].Background).ForegroundColor);

            using var stream = new MemoryStream(bytes);
            using PowerPointPresentation reopened = PowerPointPresentation.Load(stream);
            Assert.Equal("0A0B0C", reopened.Slides[0].GetBackground().Color);
            Assert.Equal("8899AA", reopened.Slides[1].GetBackground().Color);
            Assert.Empty(reopened.ValidateDocument());
        }

        [Fact]
        public void NativeWriter_WritesMasterAndMaterializedLayoutPictureBackgrounds() {
            byte[] imageBytes = PdfPngTestImages.CreateRgbPng(82, 113, 255);
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            SlideMasterPart masterPart = presentation.OpenXmlDocument
                .PresentationPart!.SlideMasterParts.First();
            SetPictureBackground(masterPart,
                masterPart.SlideMaster!.CommonSlideData!, imageBytes);
            presentation.AddSlide(P.SlideLayoutValues.Title);
            int blankIndex = presentation.GetLayoutIndex(
                P.SlideLayoutValues.Blank);
            SlideLayoutPart blankLayout = masterPart.SlideLayoutParts
                .ElementAt(blankIndex);
            SetPictureBackground(blankLayout,
                blankLayout.SlideLayout!.CommonSlideData!, imageBytes);
            presentation.AddSlide(0, blankIndex);

            LegacyPptWritePreflightReport preflight = presentation
                .AnalyzeLegacyPptWrite();
            Assert.True(preflight.CanWrite,
                string.Join(Environment.NewLine, preflight.Findings));
            byte[] binaryBytes = presentation.ToBytes(
                PowerPointFileFormat.Ppt);
            LegacyPptPresentation binary = LegacyPptPresentation.Load(
                binaryBytes);

            LegacyPptBackground masterBackground = Assert.IsType<
                LegacyPptBackground>(Assert.Single(binary.Masters).Background);
            Assert.Equal(LegacyPptBackgroundKind.Picture,
                masterBackground.Kind);
            Assert.Equal(imageBytes, masterBackground.Picture!.ImageBytes);
            Assert.True(binary.Slides[0].FollowsMasterBackground);
            LegacyPptBackground layoutBackground = Assert.IsType<
                LegacyPptBackground>(binary.Slides[1].Background);
            Assert.Equal(LegacyPptBackgroundKind.Picture,
                layoutBackground.Kind);
            Assert.Equal(imageBytes, layoutBackground.Picture!.ImageBytes);
            Assert.Equal(2U, Assert.Single(binary.BlipStoreEntries)
                .ReferenceCount);

            using var stream = new MemoryStream(binaryBytes,
                writable: false);
            using PowerPointPresentation reopened =
                PowerPointPresentation.Load(stream);
            Assert.All(reopened.Slides, slide => {
                PowerPointSlideBackground background = slide.GetBackground();
                Assert.Equal(PowerPointSlideBackgroundKind.Image,
                    background.Kind);
                Assert.Equal(imageBytes, background.ImageBytes);
            });
            Assert.Empty(reopened.ValidateDocument());
        }

        [Fact]
        public void NativeWriter_WritesNotesMasterBackground() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            NotesMasterPart notesMasterPart = presentation.OpenXmlDocument
                .PresentationPart!.NotesMasterPart!;
            notesMasterPart.NotesMaster!.CommonSlideData!.Background = new P.Background(
                new P.BackgroundProperties(
                    new A.SolidFill(new A.RgbColorModelHex { Val = "445566" })));
            presentation.AddSlide();

            byte[] bytes = presentation.ToBytes(PowerPointFileFormat.Ppt);
            LegacyPptPresentation binary = LegacyPptPresentation.Load(bytes);
            Assert.Equal("445566", Assert.IsType<LegacyPptBackground>(
                Assert.IsType<LegacyPptSpecialMaster>(binary.NotesMaster)
                    .Background).ForegroundColor);

            using var stream = new MemoryStream(bytes);
            using PowerPointPresentation reopened = PowerPointPresentation.Load(stream);
            A.SolidFill solid = Assert.IsType<A.SolidFill>(reopened.OpenXmlDocument
                .PresentationPart!.NotesMasterPart!.NotesMaster!.CommonSlideData!
                .Background!.BackgroundProperties!.GetFirstChild<A.SolidFill>());
            Assert.Equal("445566", solid.RgbColorModelHex!.Val!.Value);
            Assert.Empty(reopened.ValidateDocument());
        }

        private static void SetPictureBackground(OpenXmlPart ownerPart,
            P.CommonSlideData commonSlideData, byte[] imageBytes) {
            ImagePart imagePart = ownerPart.AddNewPart<ImagePart>("image/png");
            using (var image = new MemoryStream(imageBytes,
                       writable: false)) {
                imagePart.FeedData(image);
            }
            commonSlideData.Background = new P.Background(
                new P.BackgroundProperties(new A.BlipFill(
                    new A.Blip {
                        Embed = ownerPart.GetIdOfPart(imagePart)
                    },
                    new A.Stretch(new A.FillRectangle()))));
        }

        [Fact]
        public void OfficeArtStyleDecoder_ExposesBackgroundFillInputs() {
            OfficeArtShapeStyle style = OfficeArtShapeStyle.Decode(new[] {
                new OfficeArtProperty(0, 0x0180, 4),
                new OfficeArtProperty(1, 0x0181, 0x00030201),
                new OfficeArtProperty(2, 0x0182, 0x00008000),
                new OfficeArtProperty(3, 0x0183, 0x00060504),
                new OfficeArtProperty(4, 0x0184, 0x00004000),
                new OfficeArtProperty(5, 0x4186, 3),
                new OfficeArtProperty(6, 0x018B, unchecked((uint)(-45 * 65536))),
                new OfficeArtProperty(7, 0x018C, 25)
            });

            Assert.Equal(4U, style.FillType);
            Assert.NotNull(style.FillColor);
            Assert.Equal(0.5D, style.FillOpacity);
            Assert.NotNull(style.FillBackColor);
            Assert.Equal(0.25D, style.FillBackOpacity);
            Assert.Equal(3, style.FillBlipStoreIndex);
            Assert.Equal(-45D, style.FillAngleDegrees);
            Assert.Equal(25, style.FillFocusPercent);
        }

        [Fact]
        public void OfficeArtStyleDecoder_DecodesAndRejectsGradientStopArrays() {
            byte[] valid = CreateGradientStopArray(
                (0x00030201U, 0x00000000U),
                (0x00060504U, 0x00008000U),
                (0x00090807U, 0x00010000U));
            OfficeArtShapeStyle decoded = OfficeArtShapeStyle.Decode(new[] {
                new OfficeArtProperty(0, 0x8197, checked((uint)valid.Length),
                    valid.Length, complexData: valid)
            });

            Assert.False(decoded.IsFillGradientStopTableTruncated);
            Assert.Collection(decoded.FillGradientStops,
                stop => Assert.Equal(0D, stop.Position),
                stop => Assert.Equal(0.5D, stop.Position),
                stop => Assert.Equal(1D, stop.Position));
            Assert.Equal(0x00060504U, decoded.FillGradientStops[1].Color.Value);

            byte[] descending = CreateGradientStopArray(
                (0x00030201U, 0x00010000U),
                (0x00060504U, 0x00008000U));
            OfficeArtShapeStyle rejected = OfficeArtShapeStyle.Decode(new[] {
                new OfficeArtProperty(0, 0x8197, checked((uint)descending.Length),
                    descending.Length, complexData: descending)
            });

            Assert.True(rejected.IsFillGradientStopTableTruncated);
            Assert.Empty(rejected.FillGradientStops);
        }

        [Fact]
        public void BinaryImport_DecodesAndProjectsOfficeArtMasterBackground() {
            LegacyPptPresentation legacy = LegacyPptPresentation.Load(FixturePath);
            LegacyPptMaster mainMaster = legacy.Masters.First(master => master.IsMainMaster);
            LegacyPptBackground background = Assert.IsType<LegacyPptBackground>(
                mainMaster.Background);

            Assert.True(background.HasProjectableFill);
            Assert.NotEqual(LegacyPptBackgroundKind.Unsupported, background.Kind);
            LegacyPptImportReport report = legacy.CreateImportReport();
            Assert.True(report.BackgroundCount > 0);
            Assert.True(report.ProjectableBackgroundCount > 0);

            using PowerPointPresentation projected = PowerPointPresentation.Load(FixturePath);
            P.Background projectedBackground = Assert.IsType<P.Background>(projected
                .OpenXmlDocument.PresentationPart!.SlideMasterParts.First().SlideMaster!
                .CommonSlideData!.Background);
            Assert.NotNull(projectedBackground.BackgroundProperties);
            Assert.Empty(projected.ValidateDocument());
        }

        [Fact]
        public void BinaryCorpus_BackgroundShapesRemainTypedAndProjectable() {
            string corpus = Path.Combine(AppContext.BaseDirectory, "Documents",
                "LegacyPptCorpus");
            foreach (string path in Directory.GetFiles(corpus, "*.ppt")) {
                LegacyPptPresentation legacy = LegacyPptPresentation.Load(path);
                LegacyPptBackground[] backgrounds = legacy.Slides
                    .Select(slide => slide.Background)
                    .Concat(legacy.Masters.Select(master => master.Background))
                    .Concat(new[] {
                        legacy.NotesMaster?.Background,
                        legacy.HandoutMaster?.Background
                    })
                    .Where(background => background != null)
                    .Cast<LegacyPptBackground>()
                    .ToArray();

                Assert.NotEmpty(backgrounds);
                Assert.DoesNotContain(backgrounds,
                    background => background.Kind == LegacyPptBackgroundKind.Unsupported);
                Assert.All(backgrounds,
                    background => Assert.True(background.HasProjectableFill));
                Assert.DoesNotContain(legacy.Diagnostics,
                    diagnostic => diagnostic.Code == "PPT-BACKGROUND-PARTIAL");
            }
        }

        private static int ToEmus(int masterUnits) =>
            checked((int)Math.Round(masterUnits * 1587.5d, MidpointRounding.AwayFromZero));

        private static int ToMasterUnits(long emus) =>
            checked((int)Math.Round(emus / 1587.5d,
                MidpointRounding.AwayFromZero));

        private static byte[] CreateGradientStopArray(
            params (uint Color, uint Position)[] stops) {
            var data = new byte[checked(6 + stops.Length * 8)];
            BinaryPrimitives.WriteUInt16LittleEndian(data.AsSpan(0, 2),
                checked((ushort)stops.Length));
            BinaryPrimitives.WriteUInt16LittleEndian(data.AsSpan(2, 2),
                checked((ushort)stops.Length));
            BinaryPrimitives.WriteUInt16LittleEndian(data.AsSpan(4, 2), 8);
            for (int index = 0; index < stops.Length; index++) {
                int offset = 6 + index * 8;
                BinaryPrimitives.WriteUInt32LittleEndian(data.AsSpan(offset, 4),
                    stops[index].Color);
                BinaryPrimitives.WriteUInt32LittleEndian(data.AsSpan(offset + 4, 4),
                    stops[index].Position);
            }
            return data;
        }
    }
}
