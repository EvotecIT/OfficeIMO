using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;
using System.Linq;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_DocumentProperties_ProjectToExcelPropertiesAndSavedXlsx() {
            DateTime created = new DateTime(2026, 6, 25, 10, 30, 0, DateTimeKind.Utc);
            DateTime modified = new DateTime(2026, 6, 26, 11, 45, 0, DateTimeKind.Utc);
            DateTime reviewedAt = new DateTime(2026, 6, 26, 12, 15, 0, DateTimeKind.Utc);
            byte[] binaryPayload = { 0x00, 0x01, 0x42, 0x80, 0xff };
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFileWithDocumentPropertyStreams(
                LegacyXlsTestWorkbookBuilder.CreateMinimalWorkbookStream(),
                CreateSummaryInformationPropertySet(created, modified),
                CreateDocumentSummaryInformationPropertySet(reviewedAt, binaryPayload));
            string outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");

            try {
                using (ExcelDocument document = ExcelDocument.Load(new MemoryStream(compound))) {
                    AssertLegacyDocumentProperties(document, created, modified, reviewedAt, binaryPayload, assertModified: true);
                    Assert.Equal("7", document.BuiltinDocumentProperties.Revision);
                    document.Save(outputPath);
                }

                using ExcelDocument converted = ExcelDocument.Load(outputPath);
                Assert.False(converted.SourceFormat == ExcelFileFormat.Xls);
                AssertLegacyDocumentProperties(converted, created, modified, reviewedAt, binaryPayload, assertModified: false);
            } finally {
                TryDelete(outputPath);
            }
        }

        [Fact]
        public void LegacyXls_DocumentProperties_ReportUnsupportedCustomPropertyTypes() {
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFileWithDocumentPropertyStreams(
                LegacyXlsTestWorkbookBuilder.CreateMinimalWorkbookStream(),
                CreateSummaryInformationPropertySet(DateTime.UtcNow, DateTime.UtcNow),
                CreateDocumentSummaryInformationPropertySetWithUnsupportedCustomProperty());

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.False(result.HasImportErrors);
            LegacyXlsUnsupportedFeature feature = Assert.Single(result.UnsupportedFeatures, unsupported => unsupported.Kind == LegacyXlsUnsupportedFeatureKind.DocumentProperty);
            Assert.Equal("XLS-OLE-CUSTOM-DOCUMENT-PROPERTY-UNSUPPORTED", feature.Code);
            Assert.Equal("DocumentProperty:Custom:PropertyId:0x0002:Type:0x0048", feature.DetailCode);
            Assert.Contains("UnsupportedBlob", feature.Description);
            Assert.Contains(result.Diagnostics, diagnostic =>
                diagnostic.Severity == LegacyXlsDiagnosticSeverity.Info
                && diagnostic.Code == "XLS-OLE-CUSTOM-DOCUMENT-PROPERTY-UNSUPPORTED"
                && diagnostic.DetailCode == "DocumentProperty:Custom:PropertyId:0x0002:Type:0x0048");
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByKind[LegacyXlsUnsupportedFeatureKind.DocumentProperty]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByCode["XLS-OLE-CUSTOM-DOCUMENT-PROPERTY-UNSUPPORTED"]);
            Assert.Equal(1, result.ImportReport.UnsupportedProjectionGapCount);
            Assert.Equal(1, result.ImportReport.UnsupportedProjectionGapsByKind[LegacyXlsUnsupportedFeatureKind.DocumentProperty]);
        }

        [Fact]
        public void LegacyXls_DocumentProperties_ProjectScalarCustomPropertyVariantTypes() {
            DateTime oleDate = new DateTime(2026, 6, 26, 12, 30, 0, DateTimeKind.Local);
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFileWithDocumentPropertyStreams(
                LegacyXlsTestWorkbookBuilder.CreateMinimalWorkbookStream(),
                CreateSummaryInformationPropertySet(DateTime.UtcNow, DateTime.UtcNow),
                CreateDocumentSummaryInformationPropertySetWithScalarCustomPropertyTypes(oleDate));
            string outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");

            try {
                using (ExcelDocument document = ExcelDocument.Load(new MemoryStream(compound))) {
                    AssertScalarCustomPropertyVariantTypes(document, oleDate);
                    document.Save(outputPath);
                }

                using ExcelDocument converted = ExcelDocument.Load(outputPath);
                AssertScalarCustomPropertyVariantTypes(converted, oleDate);
            } finally {
                TryDelete(outputPath);
            }
        }

        private static void AssertLegacyDocumentProperties(ExcelDocument document, DateTime created, DateTime modified, DateTime reviewedAt, byte[] binaryPayload, bool assertModified) {
            Assert.Equal("Legacy Metadata Workbook", document.BuiltinDocumentProperties.Title);
            Assert.Equal("XLS metadata parity", document.BuiltinDocumentProperties.Subject);
            Assert.Equal("OfficeIMO Legacy Import", document.BuiltinDocumentProperties.Creator);
            Assert.Equal("xls, metadata, parity", document.BuiltinDocumentProperties.Keywords);
            Assert.Equal("OLE SummaryInformation comments", document.BuiltinDocumentProperties.Description);
            Assert.Equal("Legacy Category", document.BuiltinDocumentProperties.Category);
            Assert.Equal("Metadata Reviewer", document.BuiltinDocumentProperties.LastModifiedBy);
            AssertSameInstant(created, document.BuiltinDocumentProperties.Created);
            if (assertModified) {
                AssertSameInstant(modified, document.BuiltinDocumentProperties.Modified);
            }
            Assert.Equal("EvotecIT", document.ApplicationProperties.Company);
            Assert.Equal("Workbook Manager", document.ApplicationProperties.Manager);
            Assert.Equal("Ready", document.CustomDocumentProperties["ReleaseStatus"].Text);
            Assert.Equal(1998, document.CustomDocumentProperties["Ticket"].NumberInteger);
            Assert.Equal(98.5D, document.CustomDocumentProperties["Score"].NumberDouble);
            Assert.True(document.CustomDocumentProperties["Reviewed"].Bool);
            AssertSameInstant(reviewedAt, document.CustomDocumentProperties["ReviewedAt"].Date);
            Assert.Equal(ExcelCustomPropertyType.Binary, document.CustomDocumentProperties["BinaryPayload"].PropertyType);
            Assert.Equal(binaryPayload, document.CustomDocumentProperties["BinaryPayload"].Binary);
        }

        private static void AssertScalarCustomPropertyVariantTypes(ExcelDocument document, DateTime oleDate) {
            Assert.Equal(-12, document.CustomDocumentProperties["SignedByte"].Value);
            Assert.Equal((byte)250, document.CustomDocumentProperties["UnsignedByte"].Value);
            Assert.Equal(-32000, document.CustomDocumentProperties["SignedShort"].Value);
            Assert.Equal((ushort)65000, document.CustomDocumentProperties["UnsignedShort"].Value);
            Assert.Equal(-9000000000L, document.CustomDocumentProperties["SignedInt64"].Value);
            Assert.Equal(4000000000U, document.CustomDocumentProperties["UnsignedInt32"].Value);
            Assert.Equal(ulong.MaxValue, document.CustomDocumentProperties["UnsignedInt64"].Value);
            Assert.Equal(-2048, document.CustomDocumentProperties["VariantInt"].Value);
            Assert.Equal(3000000000U, document.CustomDocumentProperties["VariantUInt"].Value);
            Assert.Equal(12.5D, document.CustomDocumentProperties["SinglePrecision"].NumberDouble);
            Assert.Equal(1234.5678D, document.CustomDocumentProperties["Currency"].NumberDouble);
            AssertSameInstant(oleDate.ToUniversalTime(), document.CustomDocumentProperties["OleDate"].Date);
        }

        private static void AssertSameInstant(DateTime expected, DateTime? actual) {
            Assert.NotNull(actual);
            Assert.Equal(expected.ToUniversalTime(), actual.Value.ToUniversalTime());
        }

        private static byte[] CreateSummaryInformationPropertySet(DateTime created, DateTime modified) {
            var properties = new List<OleTestProperty> {
                OleTestProperty.Int16(1, 1200),
                OleTestProperty.String(2, "Legacy Metadata Workbook"),
                OleTestProperty.String(3, "XLS metadata parity"),
                OleTestProperty.String(4, "OfficeIMO Legacy Import"),
                OleTestProperty.String(5, "xls, metadata, parity"),
                OleTestProperty.String(6, "OLE SummaryInformation comments"),
                OleTestProperty.String(8, "Metadata Reviewer"),
                OleTestProperty.String(9, "7"),
                OleTestProperty.FileTime(12, created),
                OleTestProperty.FileTime(13, modified)
            };
            return CreateOlePropertySet(CreateOlePropertySection(properties));
        }

        private static byte[] CreateDocumentSummaryInformationPropertySet(DateTime reviewedAt, byte[] binaryPayload) {
            var firstSection = new List<OleTestProperty> {
                OleTestProperty.Int16(1, 1200),
                OleTestProperty.String(2, "Legacy Category"),
                OleTestProperty.String(14, "Workbook Manager"),
                OleTestProperty.String(15, "EvotecIT")
            };
            var customSection = new List<OleTestProperty> {
                OleTestProperty.Int16(1, 1200),
                OleTestProperty.Dictionary(0, new Dictionary<uint, string> {
                    [2] = "ReleaseStatus",
                    [3] = "Ticket",
                    [4] = "Score",
                    [5] = "Reviewed",
                    [6] = "ReviewedAt",
                    [7] = "BinaryPayload"
                }),
                OleTestProperty.String(2, "Ready"),
                OleTestProperty.Int32(3, 1998),
                OleTestProperty.Double(4, 98.5D),
                OleTestProperty.Boolean(5, true),
                OleTestProperty.FileTime(6, reviewedAt),
                OleTestProperty.Blob(7, binaryPayload)
            };

            return CreateOlePropertySet(CreateOlePropertySection(firstSection), CreateOlePropertySection(customSection));
        }

        private static byte[] CreateDocumentSummaryInformationPropertySetWithScalarCustomPropertyTypes(DateTime oleDate) {
            var firstSection = new List<OleTestProperty> {
                OleTestProperty.Int16(1, 1200)
            };
            var customSection = new List<OleTestProperty> {
                OleTestProperty.Int16(1, 1200),
                OleTestProperty.Dictionary(0, new Dictionary<uint, string> {
                    [2] = "SignedByte",
                    [3] = "UnsignedByte",
                    [4] = "SignedShort",
                    [5] = "UnsignedShort",
                    [6] = "SignedInt64",
                    [7] = "UnsignedInt32",
                    [8] = "UnsignedInt64",
                    [9] = "VariantInt",
                    [10] = "VariantUInt",
                    [11] = "SinglePrecision",
                    [12] = "Currency",
                    [13] = "OleDate"
                }),
                OleTestProperty.Int8(2, -12),
                OleTestProperty.UInt8(3, 250),
                OleTestProperty.Int16(4, -32000),
                OleTestProperty.UInt16(5, 65000),
                OleTestProperty.Int64(6, -9000000000L),
                OleTestProperty.UInt32(7, 4000000000U),
                OleTestProperty.UInt64(8, ulong.MaxValue),
                OleTestProperty.VariantInt(9, -2048),
                OleTestProperty.VariantUInt(10, 3000000000U),
                OleTestProperty.Float(11, 12.5F),
                OleTestProperty.Currency(12, 12345678L),
                OleTestProperty.OleDate(13, oleDate)
            };

            return CreateOlePropertySet(CreateOlePropertySection(firstSection), CreateOlePropertySection(customSection));
        }

        private static byte[] CreateDocumentSummaryInformationPropertySetWithUnsupportedCustomProperty() {
            var firstSection = new List<OleTestProperty> {
                OleTestProperty.Int16(1, 1200)
            };
            var customSection = new List<OleTestProperty> {
                OleTestProperty.Int16(1, 1200),
                OleTestProperty.Dictionary(0, new Dictionary<uint, string> {
                    [2] = "UnsupportedBlob"
                }),
                OleTestProperty.Raw(2, 0x0048, Enumerable.Repeat((byte)0x42, 16).ToArray())
            };

            return CreateOlePropertySet(CreateOlePropertySection(firstSection), CreateOlePropertySection(customSection));
        }

        private static byte[] CreateOlePropertySet(params byte[][] sections) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, 0xfffe);
            WriteUInt16(stream, 0);
            WriteUInt32(stream, 0);
            stream.Write(new byte[16], 0, 16);
            WriteUInt32(stream, checked((uint)sections.Length));

            int sectionOffset = 28 + sections.Length * 20;
            for (int i = 0; i < sections.Length; i++) {
                stream.Write(new byte[16], 0, 16);
                WriteUInt32(stream, checked((uint)sectionOffset));
                sectionOffset += sections[i].Length;
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

        private static void PadToInt32(Stream stream) {
            while (stream.Position % 4 != 0) {
                stream.WriteByte(0);
            }
        }

        private sealed class OleTestProperty {
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

            internal static OleTestProperty Int8(uint id, sbyte value) {
                return Raw(id, 0x0010, new[] { unchecked((byte)value) });
            }

            internal static OleTestProperty UInt8(uint id, byte value) {
                return Raw(id, 0x0011, new[] { value });
            }

            internal static OleTestProperty Int32(uint id, int value) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0x0003);
                WriteUInt16(stream, 0);
                WriteUInt32(stream, unchecked((uint)value));
                return new OleTestProperty(id, stream.ToArray());
            }

            internal static OleTestProperty UInt16(uint id, ushort value) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0x0012);
                WriteUInt16(stream, 0);
                WriteUInt16(stream, value);
                WriteUInt16(stream, 0);
                return new OleTestProperty(id, stream.ToArray());
            }

            internal static OleTestProperty UInt32(uint id, uint value) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0x0013);
                WriteUInt16(stream, 0);
                WriteUInt32(stream, value);
                return new OleTestProperty(id, stream.ToArray());
            }

            internal static OleTestProperty UInt64(uint id, ulong value) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0x0015);
                WriteUInt16(stream, 0);
                byte[] bytes = BitConverter.GetBytes(value);
                stream.Write(bytes, 0, bytes.Length);
                return new OleTestProperty(id, stream.ToArray());
            }

            internal static OleTestProperty Int64(uint id, long value) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0x0014);
                WriteUInt16(stream, 0);
                byte[] bytes = BitConverter.GetBytes(value);
                stream.Write(bytes, 0, bytes.Length);
                return new OleTestProperty(id, stream.ToArray());
            }

            internal static OleTestProperty VariantInt(uint id, int value) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0x0016);
                WriteUInt16(stream, 0);
                WriteUInt32(stream, unchecked((uint)value));
                return new OleTestProperty(id, stream.ToArray());
            }

            internal static OleTestProperty VariantUInt(uint id, uint value) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0x0017);
                WriteUInt16(stream, 0);
                WriteUInt32(stream, value);
                return new OleTestProperty(id, stream.ToArray());
            }

            internal static OleTestProperty Currency(uint id, long scaledValue) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0x0006);
                WriteUInt16(stream, 0);
                byte[] bytes = BitConverter.GetBytes(scaledValue);
                stream.Write(bytes, 0, bytes.Length);
                return new OleTestProperty(id, stream.ToArray());
            }

            internal static OleTestProperty Float(uint id, float value) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0x0004);
                WriteUInt16(stream, 0);
                byte[] bytes = BitConverter.GetBytes(value);
                stream.Write(bytes, 0, bytes.Length);
                return new OleTestProperty(id, stream.ToArray());
            }

            internal static OleTestProperty Double(uint id, double value) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0x0005);
                WriteUInt16(stream, 0);
                byte[] bytes = BitConverter.GetBytes(value);
                stream.Write(bytes, 0, bytes.Length);
                return new OleTestProperty(id, stream.ToArray());
            }

            internal static OleTestProperty OleDate(uint id, DateTime value) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0x0007);
                WriteUInt16(stream, 0);
                byte[] bytes = BitConverter.GetBytes(value.ToOADate());
                stream.Write(bytes, 0, bytes.Length);
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
                byte[] bytes = BitConverter.GetBytes(value.ToUniversalTime().ToFileTimeUtc());
                stream.Write(bytes, 0, bytes.Length);
                return new OleTestProperty(id, stream.ToArray());
            }

            internal static OleTestProperty Blob(uint id, byte[] value) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0x0041);
                WriteUInt16(stream, 0);
                WriteUInt32(stream, checked((uint)value.Length));
                stream.Write(value, 0, value.Length);
                PadToInt32(stream);
                return new OleTestProperty(id, stream.ToArray());
            }

            internal static OleTestProperty String(uint id, string value) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0x001f);
                WriteUInt16(stream, 0);
                WriteUInt32(stream, checked((uint)(value.Length + 1)));
                byte[] bytes = Encoding.Unicode.GetBytes(value + '\0');
                stream.Write(bytes, 0, bytes.Length);
                PadToInt32(stream);
                return new OleTestProperty(id, stream.ToArray());
            }

            internal static OleTestProperty Raw(uint id, ushort type, byte[] valueBytes) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, type);
                WriteUInt16(stream, 0);
                stream.Write(valueBytes, 0, valueBytes.Length);
                PadToInt32(stream);
                return new OleTestProperty(id, stream.ToArray());
            }

            internal static OleTestProperty Dictionary(uint id, IReadOnlyDictionary<uint, string> names) {
                using var stream = new MemoryStream();
                WriteUInt32(stream, checked((uint)names.Count));
                foreach (KeyValuePair<uint, string> name in names.OrderBy(entry => entry.Key)) {
                    WriteUInt32(stream, name.Key);
                    WriteUInt32(stream, checked((uint)(name.Value.Length + 1)));
                    byte[] bytes = Encoding.Unicode.GetBytes(name.Value + '\0');
                    stream.Write(bytes, 0, bytes.Length);
                    PadToInt32(stream);
                }

                return new OleTestProperty(id, stream.ToArray());
            }
        }
    }
}
