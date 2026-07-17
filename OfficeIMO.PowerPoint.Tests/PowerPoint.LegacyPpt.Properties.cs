using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.VariantTypes;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointLegacyPptPropertyTests {
        private const string SummaryInformationStream =
            "\u0005SummaryInformation";
        private const string DocumentSummaryInformationStream =
            "\u0005DocumentSummaryInformation";
        private static readonly Guid SummaryInformationFormatId = new(
            "F29F85E0-4FF9-1068-AB91-08002B27B3D9");

        [Fact]
        public void NativeWriter_RoundTripsOleBuiltInAndCustomProperties() {
            DateTime created = new DateTime(2025, 2, 3, 4, 5, 6,
                DateTimeKind.Utc);
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            presentation.AddSlide();
            presentation.BuiltinDocumentProperties.Title = "Binary metadata";
            presentation.BuiltinDocumentProperties.Subject = "OLE property sets";
            presentation.BuiltinDocumentProperties.Creator = "OfficeIMO";
            presentation.BuiltinDocumentProperties.Keywords = "ppt,metadata";
            presentation.BuiltinDocumentProperties.Description = "Round trip";
            presentation.BuiltinDocumentProperties.Category = "Testing";
            presentation.BuiltinDocumentProperties.Created = created;
            presentation.BuiltinDocumentProperties.Modified = created.AddHours(1);
            presentation.ApplicationProperties.Application =
                "OfficeIMO.PowerPoint";
            presentation.ApplicationProperties.TotalTime = "42";
            presentation.ApplicationProperties.Company = "Evotec";
            presentation.ApplicationProperties.Manager = "Automation";
            AddCustomTextProperty(presentation, "ReleaseStatus", "Ready");

            LegacyPptWritePreflightReport preflight = presentation
                .AnalyzeLegacyPptWrite();
            Assert.True(preflight.CanWrite,
                string.Join(Environment.NewLine, preflight.Findings));
            byte[] bytes = presentation.ToBytes(PowerPointFileFormat.Ppt);
            LegacyPptPresentation binary = LegacyPptPresentation.Load(bytes);

            IReadOnlyDictionary<string, byte[]> streams = binary.Package
                .CopyCompoundStreams();
            Assert.Contains(SummaryInformationStream, streams.Keys);
            Assert.Contains(DocumentSummaryInformationStream, streams.Keys);

            using var input = new MemoryStream(bytes);
            using PowerPointPresentation reopened =
                PowerPointPresentation.Load(input);
            Assert.Equal("Binary metadata",
                reopened.BuiltinDocumentProperties.Title);
            Assert.Equal("OLE property sets",
                reopened.BuiltinDocumentProperties.Subject);
            Assert.Equal("OfficeIMO",
                reopened.BuiltinDocumentProperties.Creator);
            Assert.Equal("ppt,metadata",
                reopened.BuiltinDocumentProperties.Keywords);
            Assert.Equal("Round trip",
                reopened.BuiltinDocumentProperties.Description);
            Assert.Equal("Testing",
                reopened.BuiltinDocumentProperties.Category);
            Assert.Equal(created,
                reopened.BuiltinDocumentProperties.Created);
            Assert.Equal(created.AddHours(1),
                reopened.BuiltinDocumentProperties.Modified);
            Assert.Equal("OfficeIMO.PowerPoint",
                reopened.ApplicationProperties.Application);
            Assert.Equal("42", reopened.ApplicationProperties.TotalTime);
            Assert.Equal("Evotec",
                reopened.ApplicationProperties.Company);
            Assert.Equal("Automation",
                reopened.ApplicationProperties.Manager);
            CustomDocumentProperty custom = Assert.Single(reopened
                .OpenXmlDocument.CustomFilePropertiesPart!.Properties!
                .Elements<CustomDocumentProperty>());
            Assert.Equal("ReleaseStatus", custom.Name!.Value);
            Assert.Equal("Ready", custom.VTLPWSTR!.Text);
            Assert.Empty(reopened.ValidateDocument());
        }

        [Fact]
        public void NativeWriter_PreservesFloatCustomPropertyVariant() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            presentation.AddSlide();
            CustomFilePropertiesPart part = presentation.OpenXmlDocument
                .AddCustomFilePropertiesPart();
            part.Properties = new Properties(new CustomDocumentProperty(
                new VTFloat("1.25")) {
                FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}",
                PropertyId = 2,
                Name = "SinglePrecision"
            });

            byte[] bytes = presentation.ToBytes(PowerPointFileFormat.Ppt);

            using var input = new MemoryStream(bytes, writable: false);
            using PowerPointPresentation reopened =
                PowerPointPresentation.Load(input);
            CustomDocumentProperty custom = Assert.Single(reopened
                .OpenXmlDocument.CustomFilePropertiesPart!.Properties!
                .Elements<CustomDocumentProperty>());
            Assert.Equal("1.25", custom.VTFloat?.Text);
            Assert.Null(custom.VTDouble);
            Assert.Empty(reopened.ValidateDocument());
        }

        [Fact]
        public void ReadOnlyImport_ProjectsOlePropertiesBeforeOpeningPackage() {
            DateTime createdAt = new DateTime(2025, 6, 7, 8, 9, 10,
                DateTimeKind.Utc);
            byte[] bytes;
            using (PowerPointPresentation created =
                   PowerPointPresentation.Create()) {
                created.AddSlide().AddTitle("Read-only binary import");
                created.BuiltinDocumentProperties.Title =
                    "Read-only metadata";
                created.BuiltinDocumentProperties.Created = createdAt;
                created.BuiltinDocumentProperties.Modified =
                    createdAt.AddMinutes(5);
                bytes = created.ToBytes(PowerPointFileFormat.Ppt);
            }

            using var input = new MemoryStream(bytes, writable: false);
            using PowerPointPresentation reopened =
                PowerPointPresentation.Load(input, new PowerPointLoadOptions {
                    AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly
                });

            Assert.Equal(OfficeIMO.Drawing.DocumentAccessMode.ReadOnly,
                reopened.AccessMode);
            Assert.Equal(PowerPointFileFormat.Ppt, reopened.SourceFormat);
            Assert.Equal("Read-only metadata",
                reopened.BuiltinDocumentProperties.Title);
            Assert.Equal(createdAt,
                reopened.BuiltinDocumentProperties.Created);
            Assert.Equal(createdAt.AddMinutes(5),
                reopened.BuiltinDocumentProperties.Modified);
            Assert.Contains(reopened.Slides[0].TextBoxes,
                textBox => textBox.Text == "Read-only binary import");
        }

        [Fact]
        public void ImportedPropertyEdits_RewriteOnlyPropertySetStreams() {
            byte[] sourceBytes;
            using (PowerPointPresentation created =
                   PowerPointPresentation.Create()) {
                created.AddSlide();
                created.BuiltinDocumentProperties.Title = "Original";
                created.BuiltinDocumentProperties.Category = "Draft";
                created.ApplicationProperties.Company = "Before";
                AddCustomTextProperty(created, "ReleaseStatus", "Pending");
                sourceBytes = created.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation original = LegacyPptPresentation.Load(
                sourceBytes);

            using var input = new MemoryStream(sourceBytes);
            using PowerPointPresentation imported =
                PowerPointPresentation.Load(input);
            imported.BuiltinDocumentProperties.Title = "Edited";
            imported.BuiltinDocumentProperties.Category = "Published";
            imported.ApplicationProperties.Company = "After";
            CustomDocumentProperty custom = Assert.Single(imported
                .OpenXmlDocument.CustomFilePropertiesPart!.Properties!
                .Elements<CustomDocumentProperty>());
            custom.VTLPWSTR!.Text = "Ready";

            LegacyPptWritePreflightReport preflight = imported
                .AnalyzeLegacyPptWrite();
            Assert.True(preflight.CanWrite,
                string.Join(Environment.NewLine, preflight.Findings));
            byte[] savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            LegacyPptPresentation saved = LegacyPptPresentation.Load(savedBytes);

            Assert.Equal(original.Package.DocumentStream,
                saved.Package.DocumentStream);
            Assert.Equal(original.Package.CurrentUserStream,
                saved.Package.CurrentUserStream);
            Assert.Equal(original.Package.UserEdits.Count,
                saved.Package.UserEdits.Count);
            IReadOnlyDictionary<string, byte[]> originalStreams = original
                .Package.CopyCompoundStreams();
            IReadOnlyDictionary<string, byte[]> savedStreams = saved.Package
                .CopyCompoundStreams();
            Assert.NotEqual(originalStreams[SummaryInformationStream],
                savedStreams[SummaryInformationStream]);
            Assert.NotEqual(originalStreams[DocumentSummaryInformationStream],
                savedStreams[DocumentSummaryInformationStream]);

            using var reopenedInput = new MemoryStream(savedBytes);
            using PowerPointPresentation reopened =
                PowerPointPresentation.Load(reopenedInput);
            Assert.Equal("Edited", reopened.BuiltinDocumentProperties.Title);
            Assert.Equal("Published",
                reopened.BuiltinDocumentProperties.Category);
            Assert.Equal("After", reopened.ApplicationProperties.Company);
            Assert.Equal("Ready", Assert.Single(reopened.OpenXmlDocument
                .CustomFilePropertiesPart!.Properties!
                .Elements<CustomDocumentProperty>()).VTLPWSTR!.Text);
            Assert.Empty(reopened.ValidateDocument());
        }

        [Fact]
        public void ImportedPropertyEdit_WithDelimiterCollisionIsPersisted() {
            byte[] sourceBytes;
            using (PowerPointPresentation created =
                   PowerPointPresentation.Create()) {
                created.AddSlide();
                created.BuiltinDocumentProperties.Creator = "a|b";
                created.BuiltinDocumentProperties.Title = "c";
                sourceBytes = created.ToBytes(PowerPointFileFormat.Ppt);
            }

            byte[] savedBytes;
            using (var input = new MemoryStream(sourceBytes))
            using (PowerPointPresentation imported =
                   PowerPointPresentation.Load(input)) {
                imported.BuiltinDocumentProperties.Creator = "a";
                imported.BuiltinDocumentProperties.Title = "b|c";
                savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }

            Assert.NotEqual(sourceBytes, savedBytes);
            using var reopenedInput = new MemoryStream(savedBytes);
            using PowerPointPresentation reopened =
                PowerPointPresentation.Load(reopenedInput);
            Assert.Equal("a", reopened.BuiltinDocumentProperties.Creator);
            Assert.Equal("b|c", reopened.BuiltinDocumentProperties.Title);
        }

        [Fact]
        public void ImportedUnknownProperty_IsPreservedForUnrelatedEditsAndBlocksMetadataLoss() {
            byte[] sourceBytes;
            using (PowerPointPresentation created =
                   PowerPointPresentation.Create()) {
                created.AddSlide().AddTitle("Property guard");
                sourceBytes = created.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation source = LegacyPptPresentation.Load(
                sourceBytes);
            byte[] unknownSummary = CreateSummaryWithUnknownProperty();
            byte[] guardedBytes = source.Package.RewriteCompoundStreams(
                new Dictionary<string, byte[]>(
                    StringComparer.OrdinalIgnoreCase) {
                    [SummaryInformationStream] = unknownSummary
                });

            using (var input = new MemoryStream(guardedBytes))
            using (PowerPointPresentation imported =
                   PowerPointPresentation.Load(input)) {
                imported.Slides[0].TextBoxes.Single(textBox =>
                    textBox.Text == "Property guard").Left += 15875L;
                LegacyPptWritePreflightReport preflight = imported
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                LegacyPptPresentation saved = LegacyPptPresentation.Load(
                    imported.ToBytes(PowerPointFileFormat.Ppt));
                Assert.Equal(unknownSummary, saved.Package
                    .CopyCompoundStreams()[SummaryInformationStream]);
            }

            using var guardedInput = new MemoryStream(guardedBytes);
            using PowerPointPresentation metadataEdit =
                PowerPointPresentation.Load(guardedInput);
            metadataEdit.BuiltinDocumentProperties.Title = "Would lose data";

            LegacyPptWritePreflightReport blocked = metadataEdit
                .AnalyzeLegacyPptWrite();
            Assert.False(blocked.CanWrite);
            Assert.Contains(blocked.Findings, finding =>
                finding.Code == "PPT-WRITE-IMPORT-LOSS");
        }

        private static void AddCustomTextProperty(
            PowerPointPresentation presentation, string name, string value) {
            CustomFilePropertiesPart part = presentation.OpenXmlDocument
                .CustomFilePropertiesPart
                ?? presentation.OpenXmlDocument.AddCustomFilePropertiesPart();
            part.Properties ??= new Properties();
            int id = part.Properties.Elements<CustomDocumentProperty>()
                .Select(property => property.PropertyId?.Value ?? 1)
                .DefaultIfEmpty(1).Max() + 1;
            part.Properties.Append(new CustomDocumentProperty(
                new VTLPWSTR(value)) {
                FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}",
                PropertyId = id,
                Name = name
            });
        }

        private static byte[] CreateSummaryWithUnknownProperty() {
            var properties = new List<(uint Id, byte[] Value)> {
                (1, CreateInt16Property(1200)),
                (2, CreateStringProperty("Guarded")),
                (99, CreateStringProperty("Opaque metadata"))
            };
            byte[] section = CreatePropertySection(properties);
            using var stream = new MemoryStream();
            WriteUInt16(stream, 0xFFFE);
            WriteUInt16(stream, 0);
            WriteUInt32(stream, 0);
            WriteBytes(stream, new byte[16]);
            WriteUInt32(stream, 1);
            WriteBytes(stream, SummaryInformationFormatId.ToByteArray());
            WriteUInt32(stream, 48);
            WriteBytes(stream, section);
            return stream.ToArray();
        }

        private static byte[] CreatePropertySection(
            IReadOnlyList<(uint Id, byte[] Value)> properties) {
            using var values = new MemoryStream();
            var offsets = new List<uint>();
            foreach ((uint _, byte[] value) in properties) {
                offsets.Add(checked((uint)(8 + properties.Count * 8
                    + values.Length)));
                WriteBytes(values, value);
                while (values.Length % 4 != 0) values.WriteByte(0);
            }
            using var stream = new MemoryStream();
            WriteUInt32(stream, checked((uint)(8 + properties.Count * 8
                + values.Length)));
            WriteUInt32(stream, checked((uint)properties.Count));
            for (int index = 0; index < properties.Count; index++) {
                WriteUInt32(stream, properties[index].Id);
                WriteUInt32(stream, offsets[index]);
            }
            WriteBytes(stream, values.ToArray());
            return stream.ToArray();
        }

        private static byte[] CreateInt16Property(short value) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, 0x0002);
            WriteUInt16(stream, 0);
            WriteUInt16(stream, unchecked((ushort)value));
            WriteUInt16(stream, 0);
            return stream.ToArray();
        }

        private static byte[] CreateStringProperty(string value) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, 0x001F);
            WriteUInt16(stream, 0);
            WriteUInt32(stream, checked((uint)(value.Length + 1)));
            WriteBytes(stream,
                System.Text.Encoding.Unicode.GetBytes(value + '\0'));
            while (stream.Length % 4 != 0) stream.WriteByte(0);
            return stream.ToArray();
        }

        private static void WriteUInt16(Stream stream, ushort value) {
            stream.WriteByte((byte)value);
            stream.WriteByte((byte)(value >> 8));
        }

        private static void WriteBytes(Stream stream, byte[] bytes) =>
            stream.Write(bytes, 0, bytes.Length);

        private static void WriteUInt32(Stream stream, uint value) {
            stream.WriteByte((byte)value);
            stream.WriteByte((byte)(value >> 8));
            stream.WriteByte((byte)(value >> 16));
            stream.WriteByte((byte)(value >> 24));
        }
    }
}
