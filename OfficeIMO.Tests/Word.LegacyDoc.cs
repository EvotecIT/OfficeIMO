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
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsDocumentPropertiesAndCustomProperties() {
            DateTime created = new DateTime(2026, 6, 29, 8, 0, 0, DateTimeKind.Utc);
            DateTime modified = new DateTime(2026, 6, 29, 9, 15, 0, DateTimeKind.Utc);
            byte[] docBytes = LegacyDocTestBuilder.CreateSimpleDocWithDocumentProperties(created, modified, "Metadata paragraph");

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            Assert.True(result.HasDocument);
            Assert.Equal(13, result.ImportReport.DocumentPropertyCount);
            Assert.Equal("Legacy DOC Metadata Title", result.Document.BuiltinDocumentProperties.Title);
            Assert.Equal("Legacy DOC metadata subject", result.Document.BuiltinDocumentProperties.Subject);
            Assert.Equal("OfficeIMO Legacy Import", result.Document.BuiltinDocumentProperties.Creator);
            Assert.Equal("doc, metadata, officeimo", result.Document.BuiltinDocumentProperties.Keywords);
            Assert.Equal("OLE SummaryInformation comments", result.Document.BuiltinDocumentProperties.Description);
            Assert.Equal("Legacy Category", result.Document.BuiltinDocumentProperties.Category);
            AssertSameInstant(created, result.Document.BuiltinDocumentProperties.Created);
            AssertSameInstant(modified, result.Document.BuiltinDocumentProperties.Modified);
            Assert.Equal("EvotecIT", result.Document.ApplicationProperties.Company);
            Assert.Equal("Document Manager", result.Document.ApplicationProperties.Manager?.Text);
            Assert.Equal("Ready", result.Document.CustomDocumentProperties["ReleaseStatus"].Text);
            Assert.True(result.Document.CustomDocumentProperties["Reviewed"].Bool);
            Assert.Equal(2003, result.Document.CustomDocumentProperties["Ticket"].NumberInteger);

            using WordDocument converted = WordDocument.Load(new MemoryStream(result.Document.SaveAsByteArray()));
            Assert.False(converted.WasLoadedFromLegacyDoc);
            Assert.Equal("Legacy DOC Metadata Title", converted.BuiltinDocumentProperties.Title);
            Assert.Equal("EvotecIT", converted.ApplicationProperties.Company);
            Assert.Equal("Ready", converted.CustomDocumentProperties["ReleaseStatus"].Text);
            Assert.True(converted.CustomDocumentProperties["Reviewed"].Bool);
            Assert.Equal(2003, converted.CustomDocumentProperties["Ticket"].NumberInteger);
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
        public void LegacyDoc_LoadLegacyDocWithReport_ImportsWordComDocFixture() {
            string docPath = GetFixtureDoc(Path.Combine("LegacyDocCorpus", "ComSimpleParagraphs.doc"));

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(docPath);

            result.EnsureNoImportErrors();
            Assert.True(result.HasDocument);
            Assert.True(result.Document.WasLoadedFromLegacyDoc);
            Assert.Equal(string.Empty, result.Document.FilePath);

            string[] paragraphs = result.Document.Paragraphs
                .Select(paragraph => paragraph.Text)
                .Where(text => !string.IsNullOrWhiteSpace(text))
                .ToArray();

            Assert.Contains("First COM paragraph", paragraphs);
            Assert.Contains("Second COM paragraph", paragraphs);
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
        public void LegacyDoc_NormalLoad_BlocksAutoSaveForLegacyDocProjection() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                File.WriteAllBytes(docPath, LegacyDocTestBuilder.CreateSimpleDoc("No autosave"));

                NotSupportedException exception = Assert.Throws<NotSupportedException>(() => WordDocument.Load(docPath, autoSave: true));

                Assert.Contains("Auto-save is not supported", exception.Message);
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using (WordDocument document = WordDocument.Create()) {
                    document.AddParagraph("Zażółć gęślą jaźń");
                    document.AddParagraph("Second plain paragraph");

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                Assert.Equal(string.Empty, reloaded.FilePath);
                string[] paragraphs = reloaded.Paragraphs
                    .Select(paragraph => paragraph.Text)
                    .Where(text => !string.IsNullOrEmpty(text))
                    .ToArray();
                Assert.Equal(new[] { "Zażółć gęślą jaźń", "Second plain paragraph" }, paragraphs);
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_BlocksFormattedRunsBeforeCreatingFile() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using WordDocument document = WordDocument.Create();
                document.AddParagraph("Formatted").SetBold();

                NotSupportedException exception = Assert.Throws<NotSupportedException>(() => document.Save(docPath));

                Assert.Contains("Run properties", exception.Message);
                Assert.False(File.Exists(docPath));
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

            internal static byte[] CreateSimpleDocWithDocumentProperties(DateTime created, DateTime modified, params string[] paragraphs) {
                string text = string.Join("\r", paragraphs) + "\r";
                byte[] wordDocumentStream = CreateWordDocumentStream(text);
                byte[] tableStream = CreateTableStream(text.Length);
                byte[] summaryInformation = CreateSummaryInformationPropertySet(created, modified);
                byte[] documentSummaryInformation = CreateDocumentSummaryInformationPropertySet();

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                    WriteStream(root, "\u0005SummaryInformation", summaryInformation);
                    WriteStream(root, "\u0005DocumentSummaryInformation", documentSummaryInformation);
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
                WriteUInt32(table, 15, 0x40000000U | ((uint)textOffset * 2U));
                WriteUInt16(table, 19, 0);
                return table;
            }

            private static byte[] CreateSummaryInformationPropertySet(DateTime created, DateTime modified) {
                var properties = new List<OleTestProperty> {
                    OleTestProperty.Int16(1, 1200),
                    OleTestProperty.String(2, "Legacy DOC Metadata Title"),
                    OleTestProperty.String(3, "Legacy DOC metadata subject"),
                    OleTestProperty.String(4, "OfficeIMO Legacy Import"),
                    OleTestProperty.String(5, "doc, metadata, officeimo"),
                    OleTestProperty.String(6, "OLE SummaryInformation comments"),
                    OleTestProperty.FileTime(12, created),
                    OleTestProperty.FileTime(13, modified)
                };

                return CreateOlePropertySet(CreateOlePropertySection(properties));
            }

            private static byte[] CreateDocumentSummaryInformationPropertySet() {
                var documentProperties = new List<OleTestProperty> {
                    OleTestProperty.Int16(1, 1200),
                    OleTestProperty.String(2, "Legacy Category"),
                    OleTestProperty.String(14, "Document Manager"),
                    OleTestProperty.String(15, "EvotecIT")
                };
                var customProperties = new List<OleTestProperty> {
                    OleTestProperty.Int16(1, 1200),
                    OleTestProperty.Dictionary(0, new Dictionary<uint, string> {
                        [2] = "ReleaseStatus",
                        [3] = "Reviewed",
                        [4] = "Ticket"
                    }),
                    OleTestProperty.String(2, "Ready"),
                    OleTestProperty.Boolean(3, true),
                    OleTestProperty.Int32(4, 2003)
                };

                return CreateOlePropertySet(CreateOlePropertySection(documentProperties), CreateOlePropertySection(customProperties));
            }

            private static byte[] CreateOlePropertySet(params byte[][] sections) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0xfffe);
                WriteUInt16(stream, 0);
                WriteUInt32(stream, 0);
                stream.Write(new byte[16], 0, 16);
                WriteUInt32(stream, checked((uint)sections.Length));

                int sectionOffset = 28 + sections.Length * 20;
                foreach (byte[] section in sections) {
                    stream.Write(new byte[16], 0, 16);
                    WriteUInt32(stream, checked((uint)sectionOffset));
                    sectionOffset += section.Length;
                }

                foreach (byte[] section in sections) {
                    stream.Write(section, 0, section.Length);
                }

                return stream.ToArray();
            }

            private static byte[] CreateOlePropertySection(IReadOnlyList<OleTestProperty> properties) {
                using var values = new MemoryStream();
                var offsets = new List<uint>(properties.Count);
                foreach (OleTestProperty property in properties) {
                    offsets.Add(checked((uint)(8 + properties.Count * 8 + values.Length)));
                    values.Write(property.ValueBytes, 0, property.ValueBytes.Length);
                    PadToInt32(values);
                }

                using var stream = new MemoryStream();
                WriteUInt32(stream, checked((uint)(8 + properties.Count * 8 + values.Length)));
                WriteUInt32(stream, checked((uint)properties.Count));
                for (int i = 0; i < properties.Count; i++) {
                    WriteUInt32(stream, properties[i].PropertyId);
                    WriteUInt32(stream, offsets[i]);
                }

                byte[] valueBytes = values.ToArray();
                stream.Write(valueBytes, 0, valueBytes.Length);
                return stream.ToArray();
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

            private static void PadToInt32(Stream stream) {
                while (stream.Position % 4 != 0) {
                    stream.WriteByte(0);
                }
            }

            private static void WriteUInt16(Stream stream, ushort value) {
                stream.WriteByte((byte)(value & 0xff));
                stream.WriteByte((byte)((value >> 8) & 0xff));
            }

            private static void WriteUInt32(Stream stream, uint value) {
                stream.WriteByte((byte)(value & 0xff));
                stream.WriteByte((byte)((value >> 8) & 0xff));
                stream.WriteByte((byte)((value >> 16) & 0xff));
                stream.WriteByte((byte)((value >> 24) & 0xff));
            }

            private static void WriteUInt64(Stream stream, ulong value) {
                WriteUInt32(stream, unchecked((uint)(value & 0xffffffffUL)));
                WriteUInt32(stream, unchecked((uint)(value >> 32)));
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

            private readonly struct OleTestProperty {
                private OleTestProperty(uint propertyId, byte[] valueBytes) {
                    PropertyId = propertyId;
                    ValueBytes = valueBytes;
                }

                internal uint PropertyId { get; }

                internal byte[] ValueBytes { get; }

                internal static OleTestProperty Int16(uint id, short value) {
                    using var stream = new MemoryStream();
                    WriteUInt16(stream, 0x0002);
                    WriteUInt16(stream, 0);
                    WriteUInt16(stream, unchecked((ushort)value));
                    WriteUInt16(stream, 0);
                    return new OleTestProperty(id, stream.ToArray());
                }

                internal static OleTestProperty Int32(uint id, int value) {
                    using var stream = new MemoryStream();
                    WriteUInt16(stream, 0x0003);
                    WriteUInt16(stream, 0);
                    WriteUInt32(stream, unchecked((uint)value));
                    return new OleTestProperty(id, stream.ToArray());
                }

                internal static OleTestProperty Boolean(uint id, bool value) {
                    using var stream = new MemoryStream();
                    WriteUInt16(stream, 0x000b);
                    WriteUInt16(stream, 0);
                    WriteUInt16(stream, value ? (ushort)0xffff : (ushort)0);
                    WriteUInt16(stream, 0);
                    return new OleTestProperty(id, stream.ToArray());
                }

                internal static OleTestProperty FileTime(uint id, DateTime value) {
                    using var stream = new MemoryStream();
                    WriteUInt16(stream, 0x0040);
                    WriteUInt16(stream, 0);
                    WriteUInt64(stream, unchecked((ulong)value.ToUniversalTime().ToFileTimeUtc()));
                    return new OleTestProperty(id, stream.ToArray());
                }

                internal static OleTestProperty String(uint id, string value) {
                    using var stream = new MemoryStream();
                    WriteUInt16(stream, 0x001f);
                    WriteUInt16(stream, 0);
                    WriteUInt32(stream, checked((uint)(value.Length + 1)));
                    byte[] bytes = System.Text.Encoding.Unicode.GetBytes(value + '\0');
                    stream.Write(bytes, 0, bytes.Length);
                    PadToInt32(stream);
                    return new OleTestProperty(id, stream.ToArray());
                }

                internal static OleTestProperty Dictionary(uint id, IReadOnlyDictionary<uint, string> names) {
                    using var stream = new MemoryStream();
                    WriteUInt32(stream, checked((uint)names.Count));
                    foreach (KeyValuePair<uint, string> name in names.OrderBy(entry => entry.Key)) {
                        WriteUInt32(stream, name.Key);
                        WriteUInt32(stream, checked((uint)(name.Value.Length + 1)));
                        byte[] bytes = System.Text.Encoding.Unicode.GetBytes(name.Value + '\0');
                        stream.Write(bytes, 0, bytes.Length);
                        PadToInt32(stream);
                    }

                    return new OleTestProperty(id, stream.ToArray());
                }
            }
        }

        private static void AssertSameInstant(DateTime expected, DateTime? actual) {
            Assert.NotNull(actual);
            Assert.Equal(expected.ToUniversalTime(), actual.Value.ToUniversalTime());
        }

        private static void DeleteIfExists(string path) {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }
}
