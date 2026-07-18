using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using OfficeIMO.PowerPoint.LegacyPpt.Write;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.Tests {
    public partial class PowerPointLegacyPptTests {
        [Fact]
        public void NativeWriter_AuthorsPpt9AutomaticNumberingAcrossRunGroups() {
            byte[] bytes;
            using (PowerPointPresentation presentation = PowerPointPresentation
                       .Create()) {
                PowerPointTextBox textBox = presentation.AddSlide(
                        P.SlideLayoutValues.Blank)
                    .AddTextBoxPoints(string.Empty, 30, 30, 320, 180);
                P.Shape shape = Assert.IsType<P.Shape>(textBox.Element);
                shape.TextBody = new P.TextBody(new A.BodyProperties(),
                    new A.ListStyle(),
                    CreateNumberedParagraph("Fourth", 0,
                        A.TextAutoNumberSchemeValues.ArabicPeriod, 4),
                    new A.Paragraph(new A.Run(new A.Text("Plain"))),
                    CreateNumberedParagraph("Third letter", 2,
                        A.TextAutoNumberSchemeValues
                            .AlphaLowerCharacterParenR, 3),
                    CreateNumberedParagraph("Roman", 0,
                        A.TextAutoNumberSchemeValues
                            .RomanUpperCharacterPeriod, 9));

                LegacyPptWritePreflightReport preflight = presentation
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                bytes = presentation.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation binary = LegacyPptPresentation.Load(bytes);
            LegacyPptShape binaryShape = Assert.Single(
                Assert.Single(binary.Slides).Shapes,
                shape => shape.Text.StartsWith("Fourth",
                    StringComparison.Ordinal));
            Assert.True(binaryShape.TextBody.HasStyle9Record);
            Assert.False(binaryShape.TextBody.IsStyle9Truncated,
                string.Join(Environment.NewLine, binary.Diagnostics));
            Assert.False(binaryShape.TextBody
                .HasUnprojectedParagraphFormatting);
            Assert.Collection(binaryShape.TextBody.ParagraphRuns,
                paragraph => AssertAutoNumber(paragraph,
                    LegacyPptAutoNumberScheme.ArabicPeriod, 4),
                paragraph => {
                    Assert.Null(paragraph.HasAutoNumber);
                    Assert.Null(paragraph.AutoNumberScheme);
                },
                paragraph => AssertAutoNumber(paragraph,
                    LegacyPptAutoNumberScheme.AlphaLowerParenRight, 3),
                paragraph => AssertAutoNumber(paragraph,
                    LegacyPptAutoNumberScheme.RomanUpperPeriod, 9));
            Assert.Contains(binaryShape.TextBody.CharacterRuns,
                run => run.Ppt9RunId == 0);
            Assert.Contains(binaryShape.TextBody.CharacterRuns,
                run => run.Ppt9RunId == 1);
            Assert.Contains(binaryShape.TextBody.CharacterRuns,
                run => run.Ppt9RunId == 2);
            Assert.Contains(binaryShape.TextBody.CharacterRuns,
                run => run.Ppt9RunId == 3);

            using var stream = new MemoryStream(bytes, writable: false);
            using PowerPointPresentation reopened = PowerPointPresentation
                .Load(stream);
            A.Paragraph[] paragraphs = Assert.IsType<P.Shape>(Assert.Single(
                    reopened.Slides[0].TextBoxes).Element).TextBody!
                .Elements<A.Paragraph>().ToArray();
            Assert.Equal(4, paragraphs.Length);
            Assert.Equal(A.TextAutoNumberSchemeValues.ArabicPeriod,
                paragraphs[0].ParagraphProperties!
                    .GetFirstChild<A.AutoNumberedBullet>()!.Type!.Value);
            Assert.Null(paragraphs[1].ParagraphProperties);
            Assert.Equal(3, paragraphs[2].ParagraphProperties!
                .GetFirstChild<A.AutoNumberedBullet>()!.StartAt!.Value);
            Assert.Equal(A.TextAutoNumberSchemeValues
                    .RomanUpperCharacterPeriod,
                paragraphs[3].ParagraphProperties!
                    .GetFirstChild<A.AutoNumberedBullet>()!.Type!.Value);
            Assert.Empty(reopened.ValidateDocument());
        }

        [Fact]
        public void NativeWriter_PreservesImplicitAutomaticNumberingContinuation() {
            byte[] bytes;
            using (PowerPointPresentation presentation = PowerPointPresentation
                       .Create()) {
                PowerPointTextBox textBox = presentation.AddSlide(
                        P.SlideLayoutValues.Blank)
                    .AddTextBoxPoints(string.Empty, 30, 30, 320, 120);
                A.Paragraph first = CreateNumberedParagraph("Third", 0,
                    A.TextAutoNumberSchemeValues
                        .AlphaLowerCharacterParenR, 3);
                A.Paragraph second = CreateNumberedParagraph("Fourth", 0,
                    A.TextAutoNumberSchemeValues
                        .AlphaLowerCharacterParenR, 1);
                second.ParagraphProperties!
                    .GetFirstChild<A.AutoNumberedBullet>()!.StartAt = null;
                Assert.IsType<P.Shape>(textBox.Element).TextBody =
                    new P.TextBody(new A.BodyProperties(), new A.ListStyle(),
                        first, second);

                bytes = presentation.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptShape shape = Assert.Single(
                Assert.Single(LegacyPptPresentation.Load(bytes).Slides)
                    .Shapes);
            Assert.Collection(shape.TextBody.ParagraphRuns,
                paragraph => AssertAutoNumber(paragraph,
                    LegacyPptAutoNumberScheme.AlphaLowerParenRight, 3),
                paragraph => AssertAutoNumber(paragraph,
                    LegacyPptAutoNumberScheme.AlphaLowerParenRight, 4));

            using var input = new MemoryStream(bytes, writable: false);
            using PowerPointPresentation reopened = PowerPointPresentation
                .Load(input);
            PowerPointParagraph[] paragraphs = reopened.Slides[0].TextBoxes
                .Single().Paragraphs.ToArray();
            Assert.Equal(3, paragraphs[0].NumberingStartAt);
            Assert.Equal(4, paragraphs[1].NumberingStartAt);
            Assert.Equal(A.TextAutoNumberSchemeValues
                    .AlphaLowerCharacterParenR,
                paragraphs[1].NumberingScheme);
        }

        [Fact]
        public void NativeWriter_AuthorsPpt9MasterAutomaticNumbering() {
            byte[] bytes;
            using (PowerPointPresentation presentation =
                       PowerPointPresentation.Create()) {
                SlideMasterPart masterPart = presentation.OpenXmlDocument
                    .PresentationPart!.SlideMasterParts.Single();
                P.BodyStyle bodyStyle = masterPart.SlideMaster!.TextStyles!
                    .BodyStyle!;
                bodyStyle.Append(
                    new A.Level1ParagraphProperties(
                        new A.AutoNumberedBullet {
                            Type = A.TextAutoNumberSchemeValues.ArabicPeriod,
                            StartAt = 7
                        }),
                    new A.Level2ParagraphProperties(
                        new A.AutoNumberedBullet {
                            Type = A.TextAutoNumberSchemeValues
                                .AlphaLowerCharacterParenR,
                            StartAt = 3
                        }));
                presentation.AddSlide(P.SlideLayoutValues.Blank);

                LegacyPptWritePreflightReport preflight = presentation
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                bytes = presentation.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation binary = LegacyPptPresentation.Load(bytes);
            LegacyPptTextMasterStyle body = Assert.Single(
                Assert.Single(binary.Masters).TextMasterStyles,
                style => style.TextType == LegacyPptTextType.Body);
            Assert.Collection(body.Levels,
                level => AssertAutoNumber(level.ParagraphProperties,
                    LegacyPptAutoNumberScheme.ArabicPeriod, 7),
                level => AssertAutoNumber(level.ParagraphProperties,
                    LegacyPptAutoNumberScheme.AlphaLowerParenRight, 3));
            Assert.False(body.IsTruncated,
                string.Join(Environment.NewLine, binary.Diagnostics));
            Assert.False(body.HasUnprojectedFormatting);

            using var input = new MemoryStream(bytes, writable: false);
            using PowerPointPresentation reopened =
                PowerPointPresentation.Load(input);
            A.TextParagraphPropertiesType[] levels = reopened.OpenXmlDocument
                .PresentationPart!.SlideMasterParts.Single().SlideMaster!
                .TextStyles!.BodyStyle!.ChildElements
                .Cast<A.TextParagraphPropertiesType>().ToArray();
            Assert.Equal(7, levels[0]
                .GetFirstChild<A.AutoNumberedBullet>()!.StartAt!.Value);
            Assert.Equal(A.TextAutoNumberSchemeValues
                    .AlphaLowerCharacterParenR,
                levels[1].GetFirstChild<A.AutoNumberedBullet>()!.Type!.Value);
            Assert.Empty(reopened.ValidateDocument());
        }

        [Fact]
        public void ImportedMasterAutomaticNumbering_CanChangeAndRemoveEntries() {
            byte[] sourceBytes;
            using (PowerPointPresentation source =
                       PowerPointPresentation.Create()) {
                P.BodyStyle bodyStyle = source.OpenXmlDocument
                    .PresentationPart!.SlideMasterParts.Single().SlideMaster!
                    .TextStyles!.BodyStyle!;
                bodyStyle.Append(
                    new A.Level1ParagraphProperties(
                        new A.AutoNumberedBullet {
                            Type = A.TextAutoNumberSchemeValues.ArabicPeriod,
                            StartAt = 2
                        }),
                    new A.Level2ParagraphProperties(
                        new A.AutoNumberedBullet {
                            Type = A.TextAutoNumberSchemeValues.ArabicPeriod,
                            StartAt = 3
                        }));
                source.AddSlide(P.SlideLayoutValues.Blank);
                sourceBytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation original = LegacyPptPresentation.Load(
                sourceBytes);

            byte[] savedBytes;
            using (var input = new MemoryStream(sourceBytes,
                       writable: false))
            using (PowerPointPresentation imported =
                       PowerPointPresentation.Load(input)) {
                A.TextParagraphPropertiesType[] levels = imported
                    .OpenXmlDocument.PresentationPart!.SlideMasterParts
                    .Single().SlideMaster!.TextStyles!.BodyStyle!
                    .ChildElements.Cast<A.TextParagraphPropertiesType>()
                    .ToArray();
                A.AutoNumberedBullet first = levels[0]
                    .GetFirstChild<A.AutoNumberedBullet>()!;
                first.Type = A.TextAutoNumberSchemeValues
                    .RomanUpperCharacterPeriod;
                first.StartAt = 5;
                levels[1].RemoveAllChildren<A.AutoNumberedBullet>();
                levels[1].Append(new A.NoBullet());

                LegacyPptWritePreflightReport preflight = imported
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation saved = LegacyPptPresentation.Load(
                savedBytes);
            LegacyPptTextMasterStyle body = Assert.Single(
                Assert.Single(saved.Masters).TextMasterStyles,
                style => style.TextType == LegacyPptTextType.Body);
            AssertAutoNumber(body.Levels[0].ParagraphProperties,
                LegacyPptAutoNumberScheme.RomanUpperPeriod, 5);
            Assert.False(body.Levels[1].ParagraphProperties.HasBullet);
            Assert.Null(body.Levels[1].ParagraphProperties.HasAutoNumber);
            Assert.Equal(original.Package.UserEdits.Count + 1,
                saved.Package.UserEdits.Count);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));
        }

        [Fact]
        public void MasterPictureBullet_AddChangeAndRemove_RoundTripsNatively() {
            byte[] sourceImage = OfficePngWriter.Encode(
                new OfficeRasterImage(3, 4,
                    OfficeColor.FromRgb(80, 100, 200)));
            byte[] sourceBytes;
            using (PowerPointPresentation source =
                   PowerPointPresentation.Create()) {
                SlideMasterPart masterPart = source.OpenXmlDocument
                    .PresentationPart!.SlideMasterParts.Single();
                ImagePart imagePart = masterPart.AddImagePart(
                    DocumentFormat.OpenXml.Packaging.ImagePartType.Png);
                using (var imageStream = new MemoryStream(sourceImage,
                           writable: false)) {
                    imagePart.FeedData(imageStream);
                }
                masterPart.SlideMaster!.TextStyles!.BodyStyle!.Append(
                    new A.Level1ParagraphProperties(
                        new A.PictureBullet(new A.Blip {
                            Embed = masterPart.GetIdOfPart(imagePart)
                        })));
                source.AddSlide(P.SlideLayoutValues.Blank);
                sourceBytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation original = LegacyPptPresentation.Load(
                sourceBytes);
            LegacyPptTextMasterStyle originalBody = Assert.Single(
                Assert.Single(original.Masters).TextMasterStyles,
                style => style.TextType == LegacyPptTextType.Body);
            Assert.Same(Assert.Single(original.PictureBullets),
                Assert.Single(originalBody.Levels).ParagraphProperties
                    .PictureBullet);

            byte[] changedImage = OfficePngWriter.Encode(
                new OfficeRasterImage(5, 2,
                    OfficeColor.FromRgb(200, 110, 30)));
            byte[] changedBytes;
            using (var input = new MemoryStream(sourceBytes,
                       writable: false))
            using (PowerPointPresentation imported =
                   PowerPointPresentation.Load(input)) {
                SlideMasterPart masterPart = imported.OpenXmlDocument
                    .PresentationPart!.SlideMasterParts.Single();
                P.BodyStyle body = masterPart.SlideMaster!.TextStyles!
                    .BodyStyle!;
                A.PictureBullet first = Assert.Single(body
                    .Descendants<A.PictureBullet>());
                string relationshipId = first.GetFirstChild<A.Blip>()!
                    .Embed!.Value!;
                ImagePart imagePart = Assert.IsType<ImagePart>(
                    masterPart.GetPartById(relationshipId));
                using (var imageStream = new MemoryStream(changedImage,
                           writable: false)) {
                    imagePart.FeedData(imageStream);
                }
                body.Append(new A.Level2ParagraphProperties(
                    new A.PictureBullet(new A.Blip {
                        Embed = relationshipId
                    })));

                LegacyPptWritePreflightReport preflight = imported
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                changedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation changed = LegacyPptPresentation.Load(
                changedBytes);
            LegacyPptPictureBullet changedBullet = Assert.Single(
                changed.PictureBullets);
            Assert.Equal(changedImage, changedBullet.ImageBytes);
            LegacyPptTextMasterStyle changedBody = Assert.Single(
                Assert.Single(changed.Masters).TextMasterStyles,
                style => style.TextType == LegacyPptTextType.Body);
            Assert.Equal(2, changedBody.Levels.Count);
            Assert.All(changedBody.Levels, level => Assert.Same(
                changedBullet, level.ParagraphProperties.PictureBullet));
            Assert.Equal(original.Package.UserEdits.Count + 1,
                changed.Package.UserEdits.Count);
            Assert.True(changed.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));

            byte[] removedBytes;
            using (var input = new MemoryStream(changedBytes,
                       writable: false))
            using (PowerPointPresentation imported =
                   PowerPointPresentation.Load(input)) {
                A.TextParagraphPropertiesType first = imported
                    .OpenXmlDocument.PresentationPart!.SlideMasterParts
                    .Single().SlideMaster!.TextStyles!.BodyStyle!
                    .ChildElements.Cast<A.TextParagraphPropertiesType>()
                    .First();
                first.RemoveAllChildren<A.PictureBullet>();
                first.Append(new A.NoBullet());

                Assert.True(imported.AnalyzeLegacyPptWrite().CanWrite);
                removedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation removed = LegacyPptPresentation.Load(
                removedBytes);
            LegacyPptTextMasterStyle removedBody = Assert.Single(
                Assert.Single(removed.Masters).TextMasterStyles,
                style => style.TextType == LegacyPptTextType.Body);
            Assert.False(removedBody.Levels[0].ParagraphProperties.HasBullet);
            Assert.Null(removedBody.Levels[0].ParagraphProperties
                .PictureBullet);
            Assert.Same(Assert.Single(removed.PictureBullets),
                removedBody.Levels[1].ParagraphProperties.PictureBullet);
            Assert.Equal(changed.Package.UserEdits.Count + 1,
                removed.Package.UserEdits.Count);
            Assert.True(removed.Package.DocumentStream.AsSpan(0,
                    changed.Package.DocumentStream.Length)
                .SequenceEqual(changed.Package.DocumentStream));
        }

        [Fact]
        public void ImportedPlainText_AddsPpt9NumberingWithAppendOnlyRewrite() {
            byte[] sourceBytes;
            using (PowerPointPresentation source = PowerPointPresentation
                       .Create()) {
                source.AddSlide(P.SlideLayoutValues.Blank)
                    .AddTextBox("Number me");
                sourceBytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation original = LegacyPptPresentation.Load(
                sourceBytes);

            byte[] savedBytes;
            using (var input = new MemoryStream(sourceBytes,
                       writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation
                       .Load(input)) {
                P.Shape shape = Assert.IsType<P.Shape>(Assert.Single(
                    imported.Slides[0].TextBoxes).Element);
                A.Paragraph paragraph = Assert.Single(shape.TextBody!
                    .Elements<A.Paragraph>());
                paragraph.ParagraphProperties = new A.ParagraphProperties(
                    new A.AutoNumberedBullet {
                        Type = A.TextAutoNumberSchemeValues.ArabicParenBoth,
                        StartAt = 7
                    });
                LegacyPptWritePreflightReport preflight = imported
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation saved = LegacyPptPresentation.Load(
                savedBytes);
            LegacyPptParagraphRun paragraphRun = Assert.Single(
                Assert.Single(Assert.Single(saved.Slides).Shapes,
                    shape => shape.Text == "Number me").TextBody
                    .ParagraphRuns);
            AssertAutoNumber(paragraphRun,
                LegacyPptAutoNumberScheme.ArabicParenBoth, 7);
            Assert.Equal(original.Package.UserEdits.Count + 1,
                saved.Package.UserEdits.Count);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));
        }

        [Fact]
        public void ImportedAutomaticNumbering_CanChangeAndRemoveEntries() {
            byte[] sourceBytes;
            using (PowerPointPresentation source = PowerPointPresentation
                       .Create()) {
                PowerPointTextBox textBox = source.AddSlide(
                        P.SlideLayoutValues.Blank)
                    .AddTextBoxPoints(string.Empty, 30, 30, 320, 120);
                Assert.IsType<P.Shape>(textBox.Element).TextBody =
                    new P.TextBody(new A.BodyProperties(), new A.ListStyle(),
                        CreateNumberedParagraph("Change", 0,
                            A.TextAutoNumberSchemeValues.ArabicPeriod, 2),
                        CreateNumberedParagraph("Remove", 0,
                            A.TextAutoNumberSchemeValues.ArabicPeriod, 3));
                sourceBytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation original = LegacyPptPresentation.Load(
                sourceBytes);

            byte[] savedBytes;
            using (var input = new MemoryStream(sourceBytes,
                       writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation
                       .Load(input)) {
                A.Paragraph[] paragraphs = Assert.IsType<P.Shape>(
                        Assert.Single(imported.Slides[0].TextBoxes).Element)
                    .TextBody!.Elements<A.Paragraph>().ToArray();
                A.AutoNumberedBullet first = paragraphs[0]
                    .ParagraphProperties!
                    .GetFirstChild<A.AutoNumberedBullet>()!;
                first.Type = A.TextAutoNumberSchemeValues
                    .AlphaUpperCharacterParenBoth;
                first.StartAt = 5;
                paragraphs[1].ParagraphProperties!
                    .RemoveAllChildren<A.AutoNumberedBullet>();
                paragraphs[1].ParagraphProperties!
                    .Append(new A.NoBullet());

                LegacyPptWritePreflightReport preflight = imported
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation saved = LegacyPptPresentation.Load(
                savedBytes);
            LegacyPptShape shape = Assert.Single(
                Assert.Single(saved.Slides).Shapes,
                candidate => candidate.Text.StartsWith("Change",
                    StringComparison.Ordinal));
            AssertAutoNumber(shape.TextBody.ParagraphRuns[0],
                LegacyPptAutoNumberScheme.AlphaUpperParenBoth, 5);
            Assert.False(shape.TextBody.ParagraphRuns[1].HasBullet);
            Assert.Null(shape.TextBody.ParagraphRuns[1].HasAutoNumber);
            Assert.Equal(original.Package.UserEdits.Count + 1,
                saved.Package.UserEdits.Count);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));
        }

        [Fact]
        public void ImportedAutomaticNumbering_RemovesLastPpt9Tag() {
            byte[] sourceBytes;
            using (PowerPointPresentation source = PowerPointPresentation
                       .Create()) {
                PowerPointTextBox textBox = source.AddSlide(
                        P.SlideLayoutValues.Blank)
                    .AddTextBoxPoints(string.Empty, 30, 30, 320, 120);
                Assert.IsType<P.Shape>(textBox.Element).TextBody =
                    new P.TextBody(new A.BodyProperties(), new A.ListStyle(),
                        CreateNumberedParagraph("Remove all", 0,
                            A.TextAutoNumberSchemeValues.ArabicPeriod, 1));
                sourceBytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation original = LegacyPptPresentation.Load(
                sourceBytes);
            Assert.True(Assert.Single(Assert.Single(original.Slides).Shapes)
                .TextBody.HasStyle9Record);

            byte[] savedBytes;
            using (var input = new MemoryStream(sourceBytes,
                       writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation
                       .Load(input)) {
                A.ParagraphProperties properties = Assert.IsType<P.Shape>(
                        Assert.Single(imported.Slides[0].TextBoxes).Element)
                    .TextBody!.Elements<A.Paragraph>().Single()
                    .ParagraphProperties!;
                properties.RemoveAllChildren<A.AutoNumberedBullet>();
                properties.Append(new A.NoBullet());
                LegacyPptWritePreflightReport preflight = imported
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation saved = LegacyPptPresentation.Load(
                savedBytes);
            LegacyPptShape shape = Assert.Single(
                Assert.Single(saved.Slides).Shapes);
            Assert.False(shape.TextBody.HasStyle9Record);
            Assert.False(Assert.Single(shape.TextBody.ParagraphRuns)
                .HasBullet);
            Assert.Equal(original.Package.UserEdits.Count + 1,
                saved.Package.UserEdits.Count);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));
        }
        private static void AssertAutoNumber(
            LegacyPptParagraphRun paragraph,
            LegacyPptAutoNumberScheme scheme, short startAt) {
            Assert.True(paragraph.HasAutoNumber);
            Assert.Equal(scheme, paragraph.AutoNumberScheme);
            Assert.Equal(startAt, paragraph.AutoNumberStartAt);
        }

        private static A.Paragraph CreateNumberedParagraph(string text,
            int level, A.TextAutoNumberSchemeValues scheme, int startAt) =>
            new(new A.ParagraphProperties(
                    new A.AutoNumberedBullet {
                        Type = scheme,
                        StartAt = startAt
                    }) {
                    Level = level
                },
                new A.Run(new A.Text(text)));

    }
}
