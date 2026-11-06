using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Core.Internal;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Theory]
        [InlineData(101)]
        [InlineData(104)]
        public void LegacyWord69_LoadAndSaveDocx_PreservesEditableHeadingRunAndTable(int fibVersion) {
            string source = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".doc");
            string destination = Path.ChangeExtension(source, ".docx");
            byte[] original = CreateWord69Fixture((ushort)fibVersion);
            File.WriteAllBytes(source, original);

            using (WordDocument imported = WordDocument.Load(source)) {
                Assert.Equal(WordFileFormat.Doc, imported.SourceFormat);
                Assert.Equal(source, imported.SourcePath);
                Assert.Null(imported.FilePath);
                Assert.False(imported.AssessLegacyDocWrite(new WordSaveOptions { AllowLoss = true }).IsSupported);
                Assert.Throws<NotSupportedException>(() => imported.Save(source, new WordSaveOptions { AllowLoss = true }));
                Assert.Equal(original, File.ReadAllBytes(source));

                imported.AddParagraph("Editable output");
                imported.Save(destination);
            }

            using WordDocument reopened = WordDocument.Load(destination);
            WordParagraph heading = Assert.Single(reopened.Paragraphs, paragraph => paragraph.Text == "Heading");
            Assert.Equal(WordParagraphStyles.Heading1, heading.Style);
            Assert.True(heading.Bold);
            Assert.Equal(150, heading.FontSize?.ToPoints());
            WordTableRow row = Assert.Single(Assert.Single(reopened.Tables).Rows);
            Assert.Equal(new[] { "Left", "Right" }, row.Cells.Select(cell => Assert.Single(cell.Paragraphs).Text));
            Assert.False(Assert.Single(row.Cells[0].Paragraphs).Italic);
            Assert.True(Assert.Single(row.Cells[1].Paragraphs).Italic);
            Assert.Contains(reopened.Paragraphs, paragraph => paragraph.Text == "Editable output");
        }

        [Fact]
        public void LegacyWord69_UnsupportedRunningStory_BlocksStrictDocxSave() {
            using var input = new MemoryStream(CreateWord69Fixture(104, hasRunningStory: true));
            using WordDocument imported = WordDocument.Load(input);
            Assert.Contains(imported.LegacyDocUnsupportedFeatures, feature => feature.Code == "doc-word69-running-stories");
            using var output = new MemoryStream();
            Assert.Throws<NotSupportedException>(() => imported.Save(output));
            Assert.Equal(0, output.Length);
        }

        [Theory]
        [InlineData(0x0004, "doc-word69-complex-unsupported")]
        [InlineData(0x0100, "doc-encrypted")]
        public void LegacyWord69_UnsupportedTextProfile_FailsBeforeProjection(int flags, string code) {
            using var input = new MemoryStream(CreateWord69Fixture(104, flags: (ushort)flags));
            var error = Assert.Throws<InvalidDataException>(() => WordDocument.Load(input));
            Assert.Contains(code, error.Message);
        }

        [Fact]
        public async Task LegacyWord69_SavePaths_RefuseToReplaceMislabeledSource() {
            string source = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".docx");
            string destination = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".docx");
            byte[] original = CreateWord69Fixture(104);
            File.WriteAllBytes(source, original);

            using WordDocument imported = WordDocument.Load(source);
            var options = new WordSaveOptions { AllowLoss = true };
            Assert.Throws<NotSupportedException>(() => imported.Save(source, options));
            Assert.Throws<NotSupportedException>(() => imported.SaveCopy(source, options));
            await Assert.ThrowsAsync<NotSupportedException>(() => imported.SaveAsync(source, options));
            await Assert.ThrowsAsync<NotSupportedException>(() => imported.SaveCopyAsync(source, options));
            Assert.Equal(original, File.ReadAllBytes(source));

            imported.SaveCopy(destination, options);
            using WordDocument reopened = WordDocument.Load(destination);
            Assert.Equal(WordFileFormat.Docx, reopened.SourceFormat);
        }
        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void LegacyWord69_NonWesternRunOrInheritedStyle_ReportsLoss(bool inheritedFont) {
            using var input = new MemoryStream(CreateWord69Fixture(104, cyrillicFont: inheritedFont ? 2 : 1));
            using WordDocument imported = WordDocument.Load(input);
            Assert.Contains(imported.LegacyDocUnsupportedFeatures, item => item.Code == "doc-word69-text-codepage-unsupported");
            using var output = new MemoryStream();
            Assert.Throws<NotSupportedException>(() => imported.Save(output));
            Assert.Equal(0, output.Length);
        }

        [Theory]
        [InlineData(65, "del", "delText")]
        [InlineData(66, "ins", "t")]
        public void LegacyWord69_RevisionRuns_RetainRevisionSemantics(int operation, string revision, string textElement) {
            using var input = new MemoryStream(CreateWord69Fixture(101, revisionOperation: (byte)operation));
            using WordDocument imported = WordDocument.Load(input);
            using var output = new MemoryStream();
            imported.Save(output, new WordSaveOptions { AllowLoss = true });
            output.Position = 0;
            using var package = new ZipArchive(output, ZipArchiveMode.Read, leaveOpen: true);
            using Stream content = package.GetEntry("word/document.xml")!.Open();
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XDocument document = XDocument.Load(content);
            XElement run = Assert.Single(document.Descendants(w + revision));
            Assert.Equal("Heading", Assert.Single(run.Descendants(w + textElement)).Value);
        }

        [Fact]
        public void LegacyWord69_DocumentProperties_ProjectTabStopAndRefuseEditProtection() {
            using var unprotected = new MemoryStream(CreateWord69Fixture(104, tabStop: 900));
            using WordDocument imported = WordDocument.Load(unprotected);
            Assert.Equal(900, imported.Settings.DefaultTabStop);

            using var protectedSource = new MemoryStream(CreateWord69Fixture(104, tabStop: 900, editProtection: true));
            var error = Assert.Throws<InvalidDataException>(() => WordDocument.Load(protectedSource));
            Assert.Contains("edit protection", error.Message);
        }
        [Theory]
        [InlineData(0, false)]
        [InlineData(1, true)]
        [InlineData(255, false)]
        public void LegacyWord69_UnresolvedCharsetWithNonWesternLanguage_ReportsLoss(int charset, bool inheritedLanguage) {
            using var input = new MemoryStream(CreateWord69Fixture(104, cyrillicLanguage: inheritedLanguage ? 2 : 1,
                defaultCharset: (byte)charset));
            using WordDocument imported = WordDocument.Load(input);
            Assert.Contains(imported.LegacyDocUnsupportedFeatures, item => item.Code == "doc-word69-text-codepage-unsupported");
        }
        private static byte[] CreateWord69Fixture(ushort version, bool hasRunningStory = false, ushort flags = 0,
            int cyrillicFont = 0, byte revisionOperation = 0, ushort tabStop = 0, bool editProtection = false,
            int cyrillicLanguage = 0, byte defaultCharset = 0) {
            const int textStart = 0x300;
            const int characterBin = 0x200;
            const int paragraphBin = 0x210;
            const int styleSheet = 0x220;
            const int characterPage = 0x400;
            const int paragraphPage = 0x600;
            byte[] text = Encoding.ASCII.GetBytes("Heading\rLeft\aRight\a\aAfter\r");
            if (cyrillicFont != 0 || cyrillicLanguage != 0) text[0] = 0xC0;
            byte[] main = new byte[0x800];
            const int fontTable = 0x270;
            if (defaultCharset != byte.MaxValue) {
                WriteInt(main, 0xD0, fontTable);
                WriteInt(main, 0xD4, 26);
                WriteWord(main, fontTable, 26);
                WriteFont(0, 'A', defaultCharset);
                WriteFont(1, 'B', 204);
                WriteFont(2, 'S', 2);
            }

            if (tabStop != 0 || editProtection) {
                const int documentProperties = 0x2A0;
                WriteInt(main, 0x150, documentProperties);
                WriteInt(main, 0x154, 84);
                WriteWord(main, documentProperties + 10, tabStop);
                if (editProtection) main[documentProperties + 7] = 0x02;
            }
            WriteWord(main, 0, 0xA5DC);
            WriteWord(main, 2, version);
            WriteWord(main, 6, 0x0409);
            WriteWord(main, 0x0A, flags);
            WriteInt(main, 0x18, textStart);
            WriteInt(main, 0x1C, textStart + text.Length + (hasRunningStory ? 1 : 0));
            WriteInt(main, 0x20, main.Length);
            WriteInt(main, 0x34, text.Length);
            WriteInt(main, 0x3C, hasRunningStory ? 1 : 0);
            text.CopyTo(main, textStart);

            WriteBin(0xB8, characterBin, characterPage);
            WriteBin(0xC0, paragraphBin, paragraphPage);
            var headingFormatting = new List<byte> { 99, 0x2C, 1, 85, 1 };
            if (cyrillicFont == 1) headingFormatting.AddRange(new byte[] { 93, 1, 0 });
            if (cyrillicLanguage == 1) headingFormatting.AddRange(new byte[] { 97, 0x19, 0x04 });
            if (revisionOperation != 0) headingFormatting.AddRange(new byte[] { revisionOperation, 1 });
            WritePage(characterPage, 1, new[] { 0, 7, 13, 18, text.Length },
                new[] { headingFormatting.ToArray(), Array.Empty<byte>(), new byte[] { 86, 1 }, Array.Empty<byte>() },
                paragraph: false);
            WritePage(paragraphPage, 7, new[] { 0, 8, 13, 19, 20, text.Length },
                new[] {
                    new byte[] { 1, 0 },
                    new byte[] { 0, 0, 24, 1 },
                    new byte[] { 0, 0, 24, 1 },
                    new byte[] { 0, 0, 24, 1, 25, 1, 190, 8, 0, 2, 0, 0, 104, 1, 208, 2 },
                    new byte[] { 0, 0 }
                }, paragraph: true);

            byte[] styleName = Encoding.ASCII.GetBytes("heading 1");
            int stylePropertyOffset = (8 + 2 + styleName.Length + 1) & ~1;
            int styleLength = stylePropertyOffset + (cyrillicFont == 2 || cyrillicLanguage == 2 ? 8 : 0);
            WriteInt(main, 0x60, styleSheet);
            WriteInt(main, 0x64, 14 + styleLength);
            WriteWord(main, styleSheet, 8);
            WriteWord(main, styleSheet + 2, 2);
            WriteWord(main, styleSheet + 4, 8);
            WriteWord(main, styleSheet + 12, (ushort)styleLength);
            WriteWord(main, styleSheet + 14, 1);
            WriteWord(main, styleSheet + 16, 0xFFF1);
            main[styleSheet + 22] = (byte)styleName.Length;
            styleName.CopyTo(main, styleSheet + 23);
            if (cyrillicFont == 2 || cyrillicLanguage == 2) {
                int styleStart = styleSheet + 14;
                WriteWord(main, styleStart + 4, 2);
                int characterProperties = styleStart + stylePropertyOffset + 2;
                WriteWord(main, characterProperties, 3);
                if (cyrillicFont == 2) {
                    main[characterProperties + 2] = 93;
                    main[characterProperties + 3] = 1;
                } else {
                    main[characterProperties + 2] = 97;
                    main[characterProperties + 3] = 0x19;
                    main[characterProperties + 4] = 0x04;
                }
            }

            return OfficeCompoundFileWriter.Write(new[] { new OfficeCompoundStream("WordDocument", main) });

            void WriteFont(int index, char name, byte charset) {
                int offset = fontTable + 2 + 8 * index;
                main[offset] = 7;
                main[offset + 4] = charset;
                main[offset + 6] = (byte)name;
            }

            void WriteBin(int fibPair, int bin, int page) {
                WriteInt(main, fibPair, bin);
                WriteInt(main, fibPair + 4, 10);
                WriteInt(main, bin, textStart);
                WriteInt(main, bin + 4, textStart + text.Length);
                WriteWord(main, bin + 8, (ushort)(page / 512));
            }

            void WritePage(int offset, int stride, int[] positions, byte[][] records, bool paragraph) {
                int cursor = 510;
                for (int i = 0; i < positions.Length; i++) WriteInt(main, offset + 4 * i, textStart + positions[i]);
                for (int i = 0; i < records.Length; i++) {
                    byte[] record = records[i];
                    if (record.Length == 0) continue;
                    cursor = (cursor - 1 - record.Length) & ~1;
                    main[offset + positions.Length * 4 + i * stride] = (byte)(cursor / 2);
                    main[offset + cursor] = (byte)(paragraph ? record.Length / 2 : record.Length);
                    record.CopyTo(main, offset + cursor + 1);
                }

                main[offset + 511] = (byte)records.Length;
            }
        }

        private static void WriteWord(byte[] target, int offset, ushort value) {
            target[offset] = (byte)value;
            target[offset + 1] = (byte)(value >> 8);
        }

        private static void WriteInt(byte[] target, int offset, int value) {
            WriteWord(target, offset, (ushort)value);
            WriteWord(target, offset + 2, (ushort)(value >> 16));
        }
    }
}
