using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointSlideReaderTests {
        private const string CraftedDeck = "crafted";

        public static IEnumerable<object[]> Decks() {
            yield return new object[] { CraftedDeck };
            string root = RepositoryTestPaths.Find();
            // PowerPoint-authored templates and Keynote exports: two independent producers.
            foreach (string path in Directory.GetFiles(Path.Combine(root, "Assets", "PowerPointTemplates"), "*.pptx")
                         .Concat(Directory.GetFiles(Path.Combine(root, "OfficeIMO.TestAssets", "Documents", "IWorkCorpus", "native-exports"), "*.pptx"))
                         .OrderBy(path => path, StringComparer.Ordinal)) {
                yield return new object[] { path.Substring(root.Length + 1).Replace('\\', '/') };
            }
        }

        // Extraction from streamed slides must produce exactly the chunks extraction from the loaded model produces.
        [Theory]
        [MemberData(nameof(Decks))]
        public void StreamedExtractionMatchesTheLoadedPresentation(string deck) {
            byte[] package = deck == CraftedDeck
                ? CreateCraftedDeck()
                : File.ReadAllBytes(Path.Combine(RepositoryTestPaths.Find(), deck));
            var variants = new[] {
                new PowerPointExtractionExtensions.PowerPointExtractOptions(),
                new PowerPointExtractionExtensions.PowerPointExtractOptions { IncludeHiddenShapes = false },
                new PowerPointExtractionExtensions.PowerPointExtractOptions { IncludeMarkdown = false },
                new PowerPointExtractionExtensions.PowerPointExtractOptions { IncludeMarkdown = false, IncludeHiddenShapes = false, IncludeNotes = false, IncludeTables = false }
            };

            int compared = 0;
            foreach (PowerPointExtractionExtensions.PowerPointExtractOptions options in variants) {
                string[] loaded;
                using (var stream = new MemoryStream(package, writable: false))
                using (PowerPointDocument presentation = PowerPointDocument.Load(stream, new PowerPointLoadOptions { AccessMode = DocumentAccessMode.ReadOnly })) {
                    loaded = presentation.ExtractMarkdownChunks(options, sourcePath: "deck.pptx").Select(Describe).ToArray();
                }
                string[] streamed;
                using (var stream = new MemoryStream(package, writable: false))
                using (PowerPointSlideReader reader = PowerPointSlideReader.Open(stream)) {
                    streamed = reader.ExtractMarkdownChunks(options, sourcePath: "deck.pptx").Select(Describe).ToArray();
                }

                Assert.Equal(loaded, streamed);
                compared += loaded.Length;
            }
            Assert.True(compared > 0);
        }

        [Fact]
        public void ReadSlidesReturnsTheAuthoredSlideContent() {
            using var stream = new MemoryStream(CreateCraftedDeck(), writable: false);
            using PowerPointSlideReader reader = PowerPointSlideReader.Open(stream);
            PowerPointSlideContent[] slides = reader.ReadSlides().ToArray();

            Assert.Equal(2, reader.SlideCount);
            Assert.Equal(new[] { 1, 2 }, slides.Select(slide => slide.Number));
            Assert.Equal(new[] { false, true }, slides.Select(slide => slide.Hidden));

            PowerPointSlideContent first = slides[0];
            Assert.Equal("Streaming Review", first.Title);
            Assert.Equal("First note" + Environment.NewLine + "Second note", first.Notes);
            // The markup-compatibility shapes and the frame with unknown graphic data are not slide shapes.
            Assert.Equal(new uint?[] { 2, 3, 4, 5, 6, 7, 8, 9, 10, 12 }, first.Shapes.Select(shape => shape.Id));
            Assert.Equal(new[] { 0, 0, 0, 0, 1, 1, 2, 0, 0, 0 }, first.Shapes.Select(shape => shape.Depth));
            Assert.Equal(new[] { true, true, false, true, true, false, false, true, true, true }, first.Shapes.Select(shape => shape.IsVisible));
            Assert.Equal(PowerPointShapeContentType.Connector, first.Shapes.Last().Kind);

            PowerPointShapeContent body = first.Shapes[1];
            Assert.Equal(PowerPointShapeContentType.TextBox, body.Kind);
            Assert.Null(body.PlaceholderType);
            PowerPointParagraphContent bullet = body.Paragraphs[0];
            Assert.Equal("•", bullet.BulletCharacter);
            Assert.Equal("  padded  link [x]", bullet.Text);
            Assert.Equal(new Uri("https://example.com/a b"), bullet.Inlines[1].Hyperlink);
            PowerPointParagraphContent numbered = body.Paragraphs[1];
            Assert.True(numbered.IsNumbered);
            Assert.Equal(1, numbered.Level);
            Assert.Equal(3, numbered.NumberingStartAt);
            Assert.Equal(PowerPointNumberingScheme.RomanUpperCharacterPeriod, numbered.NumberingScheme);
            // An automatic-numbering type outside the schema still numbers the paragraph, with no scheme.
            Assert.True(body.Paragraphs[6].IsNumbered);
            Assert.Null(body.Paragraphs[6].NumberingScheme);
            Assert.Equal(new[] { PowerPointParagraphInlineKind.Run, PowerPointParagraphInlineKind.LineBreak, PowerPointParagraphInlineKind.Field, PowerPointParagraphInlineKind.Field },
                body.Paragraphs[3].Inlines.Select(inline => inline.Kind));
            Assert.Equal(" " + Environment.NewLine + "7", body.Paragraphs[3].Text);
            Assert.Equal(new Uri("#slide-2", UriKind.Relative), body.Paragraphs[4].Inlines[0].Hyperlink);

            PowerPointTableContent table = Assert.Single(first.Tables);
            Assert.Equal("Sales table", table.Name);
            Assert.True(table.HeaderRow);
            Assert.Equal(2, table.ColumnCount);
            Assert.Equal(new[] { "Region", "Revenue", "A" + Environment.NewLine + "B", "" },
                table.Rows.SelectMany(row => row.Cells).Select(cell => cell.Text));

            PowerPointImageReference image = Assert.Single(first.Images);
            PowerPointShapeContent picture = first.Shapes.Single(shape => shape.Image != null);
            Assert.Equal("Company logo", picture.AltText);
            Assert.Equal("image/png", image.ContentType);
            Assert.EndsWith(".png", image.PartUri!.OriginalString, StringComparison.Ordinal);

            Assert.Null(slides[1].Notes);
            Assert.Equal("Second slide", Assert.Single(slides[1].Shapes).Text);
        }

        [Fact]
        public void ReadSlidesReturnsEachSlideBeforeReadingTheNext() {
            byte[] package = CreateCraftedDeck(secondSlideXml: "<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:cSld>");
            using var stream = new MemoryStream(package, writable: false);
            using PowerPointSlideReader reader = PowerPointSlideReader.Open(stream);
            using IEnumerator<PowerPointSlideContent> slides = reader.ReadSlides().GetEnumerator();

            Assert.True(slides.MoveNext());
            Assert.Equal("Streaming Review", slides.Current.Title);
            Assert.Throws<XmlException>(() => slides.MoveNext());
        }

        [Fact]
        public void ReaderHoldsTheFileOnlyUntilDisposedAndHonorsCancellation() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".pptx");
            File.WriteAllBytes(path, CreateCraftedDeck());
            try {
                using var cancellation = new CancellationTokenSource();
                PowerPointSlideReader reader = PowerPointSlideReader.Open(path);
                using (File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) { }
                using (IEnumerator<PowerPointSlideContent> slides = reader.ReadSlides(cancellationToken: cancellation.Token).GetEnumerator()) {
                    Assert.True(slides.MoveNext());
                    cancellation.Cancel();
                    Assert.Throws<OperationCanceledException>(() => slides.MoveNext());
                }
                reader.Dispose();

                Assert.Throws<ObjectDisposedException>(() => reader.ReadSlides());
                File.Delete(path);
                Assert.False(File.Exists(path));
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void OpenRejectsInputOutsideTheConfiguredLimits() {
            byte[] package = CreateCraftedDeck();
            using (var stream = new MemoryStream(package, writable: false)) {
                Assert.Throws<InvalidDataException>(() => PowerPointSlideReader.Open(stream,
                    new PowerPointLoadOptions { MaxInputBytes = package.Length - 1 }, CancellationToken.None));
            }
            using (var stream = new MemoryStream(package, writable: false)) {
                Assert.Throws<OfficePackageSecurityException>(() => PowerPointSlideReader.Open(stream,
                    new PowerPointLoadOptions { PackageSecurity = new OfficePackageSecurityOptions { MaxPartUncompressedBytes = 64 } },
                    CancellationToken.None));
            }
            using (var stream = new MemoryStream(new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 }, writable: false)) {
                Assert.Throws<NotSupportedException>(() => PowerPointSlideReader.Open(stream));
            }
        }

        [Fact]
        public void OpenRejectsAPackageWhoseMainPartIsNotAPresentation() {
            using var word = new MemoryStream();
            using (WordprocessingDocument document = WordprocessingDocument.Create(word, DocumentFormat.OpenXml.WordprocessingDocumentType.Document)) {
                document.AddMainDocumentPart().Document = new DocumentFormat.OpenXml.Wordprocessing.Document(
                    new DocumentFormat.OpenXml.Wordprocessing.Body());
            }
            word.Position = 0;
            Assert.Throws<InvalidDataException>(() => PowerPointDocument.Load(word));
            word.Position = 0;
            Assert.Throws<InvalidDataException>(() => PowerPointSlideReader.Open(word));

            using var empty = new MemoryStream();
            using (PowerPointDocument presentation = PowerPointDocument.Create(empty)) presentation.Save();
            empty.Position = 0;
            using PowerPointSlideReader reader = PowerPointSlideReader.Open(empty);
            Assert.Equal(0, reader.SlideCount);
            Assert.Empty(reader.ReadSlides());
        }

        // A notes part over the character limit is an error for the caller, not a slide without notes.
        [Fact]
        public void NotesOverThePartCharacterLimitReachTheCaller() {
            using var stream = new MemoryStream();
            using (PowerPointDocument presentation = PowerPointDocument.Create(stream)) {
                PowerPointSlide slide = presentation.AddSlide();
                slide.AddTextBox("Long notes");
                slide.Notes.Text = new string('n', 50_000);
                presentation.Save();
            }
            var options = new PowerPointLoadOptions { OpenSettings = new OfficeOpenXmlLoadSettings { MaxCharactersInPart = 20_000 } };

            stream.Position = 0;
            using (PowerPointSlideReader reader = PowerPointSlideReader.Open(stream, options, CancellationToken.None)) {
                Assert.Equal("Long notes", Assert.Single(Assert.Single(reader.ReadSlides(includeNotes: false)).Shapes).Text);
            }
            stream.Position = 0;
            using (PowerPointSlideReader reader = PowerPointSlideReader.Open(stream, options, CancellationToken.None)) {
                Assert.Throws<XmlException>(() => reader.ReadSlides().ToArray());
            }
        }

        // Cancellation requested while a slide's notes are read stops that slide instead of returning it.
        [Fact]
        public void CancellationDuringNotesStopsTheSlide() {
            using var cancellation = new CancellationTokenSource();
            using var source = new CancelOnEntryStream(CreateCraftedDeck(), "ppt/notesSlides/notesSlide1.xml", cancellation);
            using PowerPointSlideReader reader = PowerPointSlideReader.Open(source);
            source.Armed = true;
            using IEnumerator<PowerPointSlideContent> slides = reader.ReadSlides(cancellationToken: cancellation.Token).GetEnumerator();

            Assert.Throws<OperationCanceledException>(() => slides.MoveNext());
            Assert.True(source.Triggered);
        }

        // Cancels when the package first reads the stored data of the named ZIP entry.
        private sealed class CancelOnEntryStream : MemoryStream {
            private readonly long _dataStart;
            private readonly CancellationTokenSource _cancellation;

            internal CancelOnEntryStream(byte[] bytes, string entryName, CancellationTokenSource cancellation) : base(bytes, writable: false) {
                _dataStart = FindEntryData(bytes, System.Text.Encoding.ASCII.GetBytes(entryName));
                _cancellation = cancellation;
            }

            internal bool Armed { get; set; }
            internal bool Triggered { get; private set; }

            public override int Read(byte[] buffer, int offset, int count) {
                long position = Position;
                int read = base.Read(buffer, offset, count);
                Inspect(position, read);
                return read;
            }

#if NETCOREAPP
            public override int Read(Span<byte> buffer) {
                long position = Position;
                int read = base.Read(buffer);
                Inspect(position, read);
                return read;
            }
#endif

            private void Inspect(long position, int count) {
                if (Armed && !Triggered && position <= _dataStart && _dataStart < position + count) {
                    Triggered = true;
                    _cancellation.Cancel();
                }
            }

            // The data of a stored entry follows its local header: 30 bytes, then the name and the extra field.
            private static long FindEntryData(byte[] bytes, byte[] name) {
                for (int index = 0; index + 30 + name.Length <= bytes.Length; index++) {
                    if (bytes[index] != 0x50 || bytes[index + 1] != 0x4B || bytes[index + 2] != 3 || bytes[index + 3] != 4) continue;
                    int nameLength = bytes[index + 26] | bytes[index + 27] << 8;
                    int extraLength = bytes[index + 28] | bytes[index + 29] << 8;
                    if (nameLength != name.Length) continue;
                    bool match = true;
                    for (int at = 0; at < name.Length && match; at++) match = bytes[index + 30 + at] == name[at];
                    if (match) return index + 30 + nameLength + extraLength;
                }
                throw new InvalidOperationException("The entry is not in the package.");
            }
        }

        private static string Describe(PowerPointExtractChunk chunk) =>
            chunk.Id + "|" + chunk.Location.Slide + "|" + chunk.Location.BlockIndex + "|" + chunk.Text + "|" + chunk.Markdown
            + "|" + string.Join(";", chunk.Warnings ?? Array.Empty<string>());

        // A first slide written by hand with the constructs whose reading must match the model: placeholders,
        // padded and whitespace-only runs, external and slide-jump hyperlinks, breaks, fields with and without text,
        // bullets and numbering, hidden shapes and groups, markup-compatibility content, a header-row table with a
        // multi-paragraph cell, a picture, an unknown graphic frame and a connector. The second slide is hidden.
        private static byte[] CreateCraftedDeck(string? secondSlideXml = null) {
            using var stream = new MemoryStream();
            using (PowerPointDocument presentation = PowerPointDocument.Create(stream)) {
                PowerPointSlide first = presentation.AddSlide();
                first.AddTextBox("placeholder");
                first.Notes.Text = "First note" + Environment.NewLine + "Second note";
                PowerPointSlide second = presentation.AddSlide();
                second.AddTextBox("Second slide");
                second.Hide();
                presentation.Save();
            }

            using (PresentationDocument package = PresentationDocument.Open(stream, true)) {
                PresentationPart presentationPart = package.PresentationPart!;
                SlidePart[] slideParts = presentationPart.Presentation!.SlideIdList!.Elements<DocumentFormat.OpenXml.Presentation.SlideId>()
                    .Select(slideId => (SlidePart)presentationPart.GetPartById(slideId.RelationshipId!.Value!)).ToArray();
                SlidePart firstPart = slideParts[0];
                ImagePart image = firstPart.AddImagePart(ImagePartType.Png);
                using (var png = new MemoryStream(Convert.FromBase64String(
                           "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=="))) {
                    image.FeedData(png);
                }
                string imageId = firstPart.GetIdOfPart(image);
                firstPart.AddHyperlinkRelationship(new Uri("https://example.com/a b"), true, "rIdLink");
                firstPart.AddPart(slideParts[1], "rIdJump");
                Feed(firstPart, CraftedSlideXml.Replace("{image}", imageId));
                if (secondSlideXml != null) Feed(slideParts[1], secondSlideXml);
            }
            return stream.ToArray();
        }

        private static void Feed(SlidePart part, string xml) {
            using var data = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(xml));
            part.FeedData(data);
        }

        private const string CraftedSlideXml = """
            <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">
              <p:cSld><p:spTree>
                <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
                <p:grpSpPr/>
                <p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/>
                  <p:txBody><a:bodyPr/><a:p><a:r><a:t>Streaming Review</a:t></a:r></a:p></p:txBody></p:sp>
                <p:sp><p:nvSpPr><p:cNvPr id="3" name="Body"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/>
                  <p:txBody><a:bodyPr/>
                    <a:p><a:pPr lvl="0"><a:buChar char="&#8226;"/></a:pPr><a:r><a:t>  padded  </a:t></a:r><a:r><a:rPr><a:hlinkClick r:id="rIdLink"/></a:rPr><a:t>link [x]</a:t></a:r></a:p>
                    <a:p><a:pPr lvl="1"><a:buAutoNum type="romanUcPeriod" startAt="3"/></a:pPr><a:r><a:t>third</a:t></a:r></a:p>
                    <a:p><a:pPr lvl="1"><a:buAutoNum type="romanUcPeriod"/></a:pPr><a:r><a:t>fourth</a:t></a:r></a:p>
                    <a:p><a:r><a:t> </a:t></a:r><a:br/><a:fld id="{1}" type="slidenum"><a:t>7</a:t></a:fld><a:fld id="{2}" type="datetime"><a:rPr lang="en-US"/></a:fld></a:p>
                    <a:p><a:r><a:rPr><a:hlinkClick r:id="rIdJump" action="ppaction://hlinksldjump"/></a:rPr><a:t>Go to two</a:t></a:r></a:p>
                    <a:p><a:endParaRPr lang="en-US"/></a:p>
                    <a:p><a:pPr><a:buAutoNum type="notAScheme"/></a:pPr><a:r><a:t>unknown numbering</a:t></a:r></a:p>
                  </p:txBody></p:sp>
                <p:sp><p:nvSpPr><p:cNvPr id="4" name="Hidden" hidden="1"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/>
                  <p:txBody><a:bodyPr/><a:p><a:r><a:t>secret</a:t></a:r></a:p></p:txBody></p:sp>
                <p:grpSp><p:nvGrpSpPr><p:cNvPr id="5" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
                  <p:sp><p:nvSpPr><p:cNvPr id="6" name="Inner"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/>
                    <p:txBody><a:bodyPr/><a:p><a:r><a:t>inside group</a:t></a:r></a:p></p:txBody></p:sp>
                  <p:grpSp><p:nvGrpSpPr><p:cNvPr id="7" name="Hidden group" hidden="1"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
                    <p:sp><p:nvSpPr><p:cNvPr id="8" name="Deep"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/>
                      <p:txBody><a:bodyPr/><a:p><a:r><a:t>hidden by its group</a:t></a:r></a:p></p:txBody></p:sp>
                  </p:grpSp>
                </p:grpSp>
                <mc:AlternateContent><mc:Choice Requires="p14"><p:sp><p:nvSpPr><p:cNvPr id="20" name="Choice"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>alternate choice</a:t></a:r></a:p></p:txBody></p:sp></mc:Choice>
                  <mc:Fallback><p:sp><p:nvSpPr><p:cNvPr id="21" name="Fallback"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>alternate fallback</a:t></a:r></a:p></p:txBody></p:sp></mc:Fallback></mc:AlternateContent>
                <p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="9" name="Sales table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></p:xfrm>
                  <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1"/><a:tblGrid><a:gridCol w="50"/><a:gridCol w="50"/></a:tblGrid>
                    <a:tr h="10"><a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>Region</a:t></a:r></a:p></a:txBody></a:tc><a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>Revenue</a:t></a:r></a:p></a:txBody></a:tc></a:tr>
                    <a:tr h="10"><a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>A</a:t></a:r></a:p><a:p><a:r><a:t>B</a:t></a:r></a:p></a:txBody></a:tc><a:tc><a:txBody><a:bodyPr/><a:p/></a:txBody></a:tc></a:tr>
                  </a:tbl></a:graphicData></a:graphic></p:graphicFrame>
                <p:pic><p:nvPicPr><p:cNvPr id="10" name="Logo" descr="Company logo"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="{image}"/></p:blipFill><p:spPr/></p:pic>
                <p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="11" name="Unknown"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></p:xfrm>
                  <a:graphic><a:graphicData uri="urn:example"><x:item xmlns:x="urn:example">ignored</x:item></a:graphicData></a:graphic></p:graphicFrame>
                <p:cxnSp><p:nvCxnSpPr><p:cNvPr id="12" name="Connector"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr/></p:cxnSp>
              </p:spTree></p:cSld>
            </p:sld>
            """;
    }
}
