using OfficeIMO.Word;
using OfficeIMO.Word.LegacyDoc;
using OpenMcdf;
using Xunit;
using Version = OpenMcdf.Version;
using StorageModeFlags = OpenMcdf.StorageModeFlags;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsBuiltInParagraphStyles() {
            byte[] docBytes = LegacyDocParagraphStyleFixture.CreateDocWithHeadingStyles();

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            WordParagraph[] paragraphs = result.Document.Paragraphs
                .Where(paragraph => !string.IsNullOrWhiteSpace(paragraph.Text))
                .ToArray();

            Assert.Equal(3, paragraphs.Length);
            Assert.Equal("Heading One", paragraphs[0].Text);
            Assert.Equal(WordParagraphStyles.Heading1, paragraphs[0].Style);
            Assert.Equal("Heading Two", paragraphs[1].Text);
            Assert.Equal(WordParagraphStyles.Heading2, paragraphs[1].Style);
            Assert.Equal("Body", paragraphs[2].Text);
            Assert.Null(paragraphs[2].Style);

            using WordDocument converted = WordDocument.Load(new MemoryStream(result.Document.SaveAsByteArray()));
            WordParagraph[] convertedParagraphs = converted.Paragraphs
                .Where(paragraph => !string.IsNullOrWhiteSpace(paragraph.Text))
                .ToArray();

            Assert.Equal(WordParagraphStyles.Heading1, convertedParagraphs[0].Style);
            Assert.Equal(WordParagraphStyles.Heading2, convertedParagraphs[1].Style);
            Assert.Null(convertedParagraphs[2].Style);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsCustomParagraphStyleFromStyleSheet() {
            byte[] docBytes = LegacyDocParagraphStyleFixture.CreateDocWithCustomParagraphStyle();

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            WordParagraph[] paragraphs = result.Document.Paragraphs
                .Where(paragraph => !string.IsNullOrWhiteSpace(paragraph.Text))
                .ToArray();

            Assert.Equal(2, paragraphs.Length);
            Assert.Equal("Styled Custom", paragraphs[0].Text);
            Assert.Equal(WordParagraphStyles.Custom, paragraphs[0].Style);
            Assert.Equal("LegacyDocCustomBody", paragraphs[0].StyleId);
            Assert.Equal("Body", paragraphs[1].Text);
            Assert.Null(paragraphs[1].Style);

            using WordDocument converted = WordDocument.Load(new MemoryStream(result.Document.SaveAsByteArray()));
            WordParagraph convertedParagraph = converted.Paragraphs
                .First(paragraph => paragraph.Text == "Styled Custom");
            Assert.Equal(WordParagraphStyles.Custom, convertedParagraph.Style);
            Assert.Equal("LegacyDocCustomBody", convertedParagraph.StyleId);

            DocumentFormat.OpenXml.Wordprocessing.Style? customStyle = converted._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .OfType<DocumentFormat.OpenXml.Wordprocessing.Style>()
                .FirstOrDefault(style => style.StyleId?.Value == "LegacyDocCustomBody");
            Assert.NotNull(customStyle);
            Assert.Equal("Custom Body", customStyle!.StyleName?.Val?.Value);
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocParagraphStylesAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using (WordDocument document = WordDocument.Create()) {
                    document.AddParagraph("Heading One").SetStyle(WordParagraphStyles.Heading1);
                    document.AddParagraph("Heading Two").SetStyle(WordParagraphStyles.Heading2);
                    document.AddParagraph("Body");

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);
                WordParagraph[] paragraphs = reloaded.Paragraphs
                    .Where(paragraph => !string.IsNullOrWhiteSpace(paragraph.Text))
                    .ToArray();

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                Assert.Equal(3, paragraphs.Length);
                Assert.Equal("Heading One", paragraphs[0].Text);
                Assert.Equal(WordParagraphStyles.Heading1, paragraphs[0].Style);
                Assert.Equal("Heading Two", paragraphs[1].Text);
                Assert.Equal(WordParagraphStyles.Heading2, paragraphs[1].Style);
                Assert.Equal("Body", paragraphs[2].Text);
                Assert.Null(paragraphs[2].Style);
            } finally {
                if (File.Exists(docPath)) {
                    File.Delete(docPath);
                }
            }
        }

        private static class LegacyDocParagraphStyleFixture {
            private const int FibLength = 0x1AA;
            private const int TextOffset = 0x200;
            private const int PapxFkpOffset = 0x400;
            private const int OleSectorSize = 512;
            private const int StyleSheetOffset = 64;
            private const ushort SprmPIstd = 0x4600;
            private const ushort CustomStyleIndex = 10;

            internal static byte[] CreateDocWithHeadingStyles() {
                const string text = "Heading One\rHeading Two\rBody\r";
                byte[] wordDocumentStream = CreateWordDocumentStream(text);
                byte[] tableStream = CreateTableStream(text.Length);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateDocWithCustomParagraphStyle() {
                const string text = "Styled Custom\rBody\r";
                byte[] styleSheet = CreateStyleSheet(new Dictionary<ushort, string> {
                    [CustomStyleIndex] = "Custom Body"
                });
                byte[] wordDocumentStream = CreateWordDocumentStream(
                    text,
                    new Dictionary<int, byte[]> {
                        [0] = CreateParagraphStylePapx(CustomStyleIndex)
                    },
                    styleSheet.Length);
                byte[] tableStream = CreateTableStream(text.Length, styleSheet);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            private static byte[] CreateWordDocumentStream(string text) {
                return CreateWordDocumentStream(
                    text,
                    new Dictionary<int, byte[]> {
                        [0] = CreateParagraphStylePapx(1),
                        [1] = CreateParagraphStyleSprmPapx(2)
                    },
                    styleSheetLength: 0);
            }

            private static byte[] CreateWordDocumentStream(string text, IReadOnlyDictionary<int, byte[]> papxByParagraphIndex, int styleSheetLength) {
                byte[] textBytes = System.Text.Encoding.Unicode.GetBytes(text);
                var stream = new byte[Math.Max(PapxFkpOffset + OleSectorSize, TextOffset + textBytes.Length)];
                WriteUInt16(stream, 0x00, 0xA5EC);
                WriteUInt16(stream, 0x02, 0x00D9);
                WriteUInt16(stream, 0x0A, 0x0200);
                WriteInt32(stream, 0x4C, text.Length);
                if (styleSheetLength > 0) {
                    WriteInt32(stream, 0xA2, StyleSheetOffset);
                    WriteInt32(stream, 0xA6, styleSheetLength);
                }

                WriteInt32(stream, 0x102, 21);
                WriteInt32(stream, 0x106, 12);
                WriteInt32(stream, 0x1A2, 0);
                WriteInt32(stream, 0x1A6, 21);
                Buffer.BlockCopy(textBytes, 0, stream, TextOffset, textBytes.Length);

                WritePapxFkp(stream, CreateParagraphPositions(text), papxByParagraphIndex);

                if (stream.Length < FibLength) {
                    Array.Resize(ref stream, FibLength);
                }

                return stream;
            }

            private static byte[] CreateTableStream(int characterCount, byte[]? styleSheet = null) {
                int length = styleSheet == null ? 33 : Math.Max(33, StyleSheetOffset + styleSheet.Length);
                var table = new byte[length];
                table[0] = 0x02;
                WriteInt32(table, 1, 16);
                WriteInt32(table, 5, 0);
                WriteInt32(table, 9, characterCount);
                WriteUInt16(table, 13, 0);
                WriteUInt32(table, 15, TextOffset);
                WriteUInt16(table, 19, 0);

                int papxPlcOffset = 21;
                WriteInt32(table, papxPlcOffset, TextOffset);
                WriteInt32(table, papxPlcOffset + 4, TextOffset + (characterCount * 2));
                WriteInt32(table, papxPlcOffset + 8, PapxFkpOffset / OleSectorSize);
                if (styleSheet != null) {
                    Buffer.BlockCopy(styleSheet, 0, table, StyleSheetOffset, styleSheet.Length);
                }

                return table;
            }

            private static int[] CreateParagraphPositions(string text) {
                var positions = new List<int> { TextOffset };
                int characterOffset = 0;
                foreach (char character in text) {
                    characterOffset++;
                    if (character == '\r') {
                        positions.Add(TextOffset + (characterOffset * 2));
                    }
                }

                return positions.ToArray();
            }

            private static byte[] CreateStyleSheet(IReadOnlyDictionary<ushort, string> styleNamesByIndex) {
                ushort cstd = checked((ushort)(styleNamesByIndex.Keys.Max() + 1));
                var bytes = new List<byte>();
                WriteUInt16(bytes, 18);
                WriteUInt16(bytes, cstd);
                WriteUInt16(bytes, 10);
                for (int i = 0; i < 7; i++) {
                    WriteUInt16(bytes, 0);
                }

                for (ushort index = 0; index < cstd; index++) {
                    if (!styleNamesByIndex.TryGetValue(index, out string? name)) {
                        WriteUInt16(bytes, 0);
                        continue;
                    }

                    byte[] std = CreateParagraphStyleDefinition(name);
                    WriteUInt16(bytes, checked((ushort)std.Length));
                    bytes.AddRange(std);
                    if (bytes.Count % 2 != 0) {
                        bytes.Add(0);
                    }
                }

                return bytes.ToArray();
            }

            private static byte[] CreateParagraphStyleDefinition(string name) {
                byte[] nameBytes = System.Text.Encoding.Unicode.GetBytes(name);
                var bytes = new List<byte>();
                WriteUInt16(bytes, 0x0FFE);
                WriteUInt16(bytes, 0x0001);
                WriteUInt16(bytes, 0);
                WriteUInt16(bytes, 0);
                WriteUInt16(bytes, 0);
                WriteUInt16(bytes, checked((ushort)name.Length));
                bytes.AddRange(nameBytes);
                WriteUInt16(bytes, 0);
                return bytes.ToArray();
            }

            private static void WritePapxFkp(byte[] stream, int[] fileParagraphPositions, IReadOnlyDictionary<int, byte[]> papxByParagraphIndex) {
                const int bxLength = 13;
                int paragraphCount = fileParagraphPositions.Length - 1;
                for (int i = 0; i < fileParagraphPositions.Length; i++) {
                    WriteInt32(stream, PapxFkpOffset + (i * 4), fileParagraphPositions[i]);
                }

                int rgbxOffset = PapxFkpOffset + (fileParagraphPositions.Length * 4);
                int papxOffset = 0x180;
                for (int i = 0; i < paragraphCount; i++) {
                    if (!papxByParagraphIndex.TryGetValue(i, out byte[]? papx)) {
                        continue;
                    }

                    papxOffset = AlignToEven(papxOffset);
                    stream[rgbxOffset + (i * bxLength)] = checked((byte)(papxOffset / 2));
                    Buffer.BlockCopy(papx, 0, stream, PapxFkpOffset + papxOffset, papx.Length);
                    papxOffset += papx.Length;
                }

                stream[PapxFkpOffset + OleSectorSize - 1] = checked((byte)paragraphCount);
            }

            private static byte[] CreateParagraphStylePapx(ushort styleIndex) {
                return CreateParagraphPropertiesPapx(styleIndex);
            }

            private static byte[] CreateParagraphStyleSprmPapx(ushort styleIndex) {
                return CreateParagraphPropertiesPapx(0, CreateParagraphSprm(SprmPIstd, (byte)(styleIndex & 0xFF), (byte)(styleIndex >> 8)));
            }

            private static byte[] CreateParagraphPropertiesPapx(ushort baseStyleIndex, params byte[][] sprms) {
                var grpprl = new List<byte> {
                    (byte)(baseStyleIndex & 0xFF),
                    (byte)(baseStyleIndex >> 8)
                };

                foreach (byte[] sprm in sprms) {
                    grpprl.AddRange(sprm);
                }

                if (grpprl.Count % 2 != 0) {
                    grpprl.Add(0);
                }

                var papx = new byte[grpprl.Count + 2];
                papx[0] = 0;
                papx[1] = checked((byte)(grpprl.Count / 2));
                grpprl.CopyTo(papx, 2);
                return papx;
            }

            private static byte[] CreateParagraphSprm(ushort sprm, params byte[] operand) {
                var bytes = new byte[2 + operand.Length];
                WriteUInt16(bytes, 0, sprm);
                Buffer.BlockCopy(operand, 0, bytes, 2, operand.Length);
                return bytes;
            }

            private static void WriteStream(RootStorage root, string name, byte[] bytes) {
                using CfbStream stream = root.CreateStream(name);
                stream.Write(bytes, 0, bytes.Length);
            }

            private static int AlignToEven(int value) {
                return value % 2 == 0 ? value : value + 1;
            }

            private static void WriteUInt16(byte[] bytes, int offset, ushort value) {
                bytes[offset] = (byte)(value & 0xFF);
                bytes[offset + 1] = (byte)(value >> 8);
            }

            private static void WriteUInt16(List<byte> bytes, int value) {
                bytes.Add((byte)(value & 0xFF));
                bytes.Add((byte)(value >> 8));
            }

            private static void WriteInt32(byte[] bytes, int offset, int value) {
                bytes[offset] = (byte)value;
                bytes[offset + 1] = (byte)(value >> 8);
                bytes[offset + 2] = (byte)(value >> 16);
                bytes[offset + 3] = (byte)(value >> 24);
            }

            private static void WriteUInt32(byte[] bytes, int offset, uint value) {
                bytes[offset] = (byte)value;
                bytes[offset + 1] = (byte)(value >> 8);
                bytes[offset + 2] = (byte)(value >> 16);
                bytes[offset + 3] = (byte)(value >> 24);
            }
        }
    }
}
