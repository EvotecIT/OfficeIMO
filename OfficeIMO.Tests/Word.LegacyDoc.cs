using OfficeIMO.Word;
using OfficeIMO.Word.LegacyDoc;
using OfficeIMO.Word.LegacyDoc.Diagnostics;
using OpenMcdf;
using Xunit;
using Version = OpenMcdf.Version;
using StorageModeFlags = OpenMcdf.StorageModeFlags;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsPlainTextParagraphs() {
            byte[] docBytes = LegacyDocTestBuilder.CreateSimpleDoc("First paragraph", "Second paragraph");

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            Assert.True(result.HasDocument);
            Assert.Equal(2, result.ImportReport.ParagraphCount);
            Assert.Equal(2, result.Document.Paragraphs.Count);
            Assert.Equal("First paragraph", result.Document.Paragraphs[0].Text);
            Assert.Equal("Second paragraph", result.Document.Paragraphs[1].Text);
            Assert.True(result.Document.WasLoadedFromLegacyDoc);
            Assert.Equal(string.Empty, result.Document.FilePath);

            using WordDocument reloaded = WordDocument.Load(new MemoryStream(result.Document.SaveAsByteArray()));
            Assert.Equal("First paragraph", reloaded.Paragraphs[0].Text);
            Assert.Equal("Second paragraph", reloaded.Paragraphs[1].Text);
        }

        [Fact]
        public void LegacyDoc_NormalLoad_RoutesOleDocIntoProjectedWordDocument() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                File.WriteAllBytes(docPath, LegacyDocTestBuilder.CreateSimpleDoc("Normal load"));

                using WordDocument document = WordDocument.Load(docPath);

                Assert.True(document.WasLoadedFromLegacyDoc);
                Assert.Equal(string.Empty, document.FilePath);
                WordParagraph paragraph = Assert.Single(document.Paragraphs);
                Assert.Equal("Normal load", paragraph.Text);
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ReportsMissingWordDocumentStream() {
            byte[] docBytes = LegacyDocTestBuilder.CreateCompoundWithoutWordDocumentStream();

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            Assert.False(result.HasDocument);
            Assert.True(result.HasImportErrors);
            LegacyDocImportDiagnostic diagnostic = Assert.Single(result.Diagnostics);
            Assert.Equal("DOC-WORDDOCUMENT-MISSING", diagnostic.Code);
            Assert.Equal(LegacyDocDiagnosticSeverity.Error, diagnostic.Severity);
        }

        [Fact]
        public void LegacyDoc_NormalLoad_BlocksAutoSaveUntilNativeDocSaveExists() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                File.WriteAllBytes(docPath, LegacyDocTestBuilder.CreateSimpleDoc("No autosave"));

                NotSupportedException exception = Assert.Throws<NotSupportedException>(() => WordDocument.Load(docPath, autoSave: true));

                Assert.Contains("Auto-save is not supported", exception.Message);
            } finally {
                DeleteIfExists(docPath);
            }
        }

        private static class LegacyDocTestBuilder {
            internal static byte[] CreateSimpleDoc(params string[] paragraphs) {
                string text = string.Join("\r", paragraphs) + "\r";
                byte[] wordDocumentStream = CreateWordDocumentStream(text);
                byte[] tableStream = CreateTableStream(text.Length);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateCompoundWithoutWordDocumentStream() {
                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "NotWordDocument", new byte[] { 1, 2, 3 });
                }

                return package.ToArray();
            }

            private static byte[] CreateWordDocumentStream(string text) {
                const int fibLength = 0x1AA;
                const int textOffset = 0x200;
                byte[] textBytes = EncodeWindows1252(text);
                var stream = new byte[textOffset + textBytes.Length];
                WriteUInt16(stream, 0x00, 0xA5EC);
                WriteUInt16(stream, 0x02, 0x00D9);
                WriteUInt16(stream, 0x0A, 0x0200);
                WriteInt32(stream, 0x4C, text.Length);
                WriteInt32(stream, 0x1A2, 0);
                WriteInt32(stream, 0x1A6, 21);
                Buffer.BlockCopy(textBytes, 0, stream, textOffset, textBytes.Length);
                if (stream.Length < fibLength) {
                    Array.Resize(ref stream, fibLength);
                }

                return stream;
            }

            private static byte[] CreateTableStream(int characterCount) {
                const int textOffset = 0x200;
                var table = new byte[21];
                table[0] = 0x02;
                WriteInt32(table, 1, 16);
                WriteInt32(table, 5, 0);
                WriteInt32(table, 9, characterCount);
                WriteUInt16(table, 13, 0);
                WriteUInt32(table, 15, 0x40000000U | textOffset);
                WriteUInt16(table, 19, 0);
                return table;
            }

            private static void WriteStream(RootStorage root, string name, byte[] bytes) {
                using CfbStream stream = root.CreateStream(name);
                stream.Write(bytes, 0, bytes.Length);
            }

            private static byte[] EncodeWindows1252(string text) {
                var bytes = new byte[text.Length];
                for (int i = 0; i < text.Length; i++) {
                    char character = text[i];
                    bytes[i] = character <= 0x7F || (character >= 0xA0 && character <= 0xFF)
                        ? (byte)character
                        : (byte)'?';
                }

                return bytes;
            }

            private static void WriteUInt16(byte[] bytes, int offset, ushort value) {
                bytes[offset] = (byte)value;
                bytes[offset + 1] = (byte)(value >> 8);
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

        private static void DeleteIfExists(string path) {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }
}
