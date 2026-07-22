using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using System.IO.Compression;
using System.Threading.Tasks;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.Tests {
    public partial class PowerPointLegacyPptTests {
        [Fact]
        public void CompoundStorageValidation_BoundsOleAndVbaLogicalExpansion() {
            var options = new LegacyPptImportOptions();
            byte[] oleStorage = CreateOleTestStorage("Bounded import OLE");
            foreach (string streamName in new[] {
                         "\u0001Ole10Native", "CONTENTS"
                     }) {
                int entry = FindCompoundDirectoryEntry(oleStorage,
                    streamName);
                WriteCompoundUInt64(oleStorage, entry + 120,
                    checked((ulong)oleStorage.Length));
            }

            Assert.False(LegacyPptCompoundStorageValidator.TryRead(
                oleStorage, options, out _, out string? oleReason));
            Assert.Contains("Compound stream bytes exceed", oleReason,
                StringComparison.OrdinalIgnoreCase);

            byte[] vbaStorage = CreateVbaTestProject("BoundedModule",
                "Sub Main(): End Sub");
            foreach (string streamName in new[] {
                         "dir", "_VBA_PROJECT"
                     }) {
                int entry = FindCompoundDirectoryEntry(vbaStorage,
                    streamName);
                WriteCompoundUInt64(vbaStorage, entry + 120,
                    checked((ulong)vbaStorage.Length));
            }

            Assert.False(LegacyPptVbaProjectCodec.IsValidProject(
                vbaStorage, options, out string? vbaReason));
            Assert.Contains("Compound stream bytes exceed", vbaReason,
                StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public async Task PresentationFacade_EnforcesPackageSecurityPolicies() {
            byte[] packageBytes;
            using (PowerPointPresentation source =
                   PowerPointPresentation.Create()) {
                source.AddSlide().AddTextBox("Security policy");
                packageBytes = source.ToBytes();
            }
            using (var editable = new MemoryStream()) {
                editable.Write(packageBytes, 0, packageBytes.Length);
                editable.Position = 0;
                using (PresentationDocument document =
                       PresentationDocument.Open(editable, true)) {
                    document.PresentationPart!.AddExternalRelationship(
                        "urn:officeimo:test", new Uri(
                            "https://example.test/presentation"),
                        "rSecurityExternal");
                }
                packageBytes = editable.ToArray();
            }
            var loadOptions = new PowerPointLoadOptions {
                PackageSecurity = OfficePackageSecurityOptions
                    .UntrustedDefaults
            };

            using (var input = new MemoryStream(packageBytes,
                       writable: false)) {
                OfficePackageSecurityException exception = Assert.Throws<
                    OfficePackageSecurityException>(() =>
                        PowerPointPresentation.Load(input, loadOptions));
                Assert.Equal(OfficePackageSecurityRule
                    .ExternalRelationships, exception.Rule);
            }
            using (var input = new MemoryStream(packageBytes,
                       writable: false)) {
                OfficePackageSecurityException exception = await Assert
                    .ThrowsAsync<OfficePackageSecurityException>(() =>
                        PowerPointPresentation.LoadAsync(input,
                            loadOptions));
                Assert.Equal(OfficePackageSecurityRule
                    .ExternalRelationships, exception.Rule);
            }

            string path = Path.Combine(Path.GetTempPath(),
                Guid.NewGuid() + ".pptx");
            try {
                File.WriteAllBytes(path, packageBytes);
                OfficePackageSecurityException exception = Assert.Throws<
                    OfficePackageSecurityException>(() =>
                        PowerPointPresentation.Load(path, loadOptions));
                Assert.Equal(OfficePackageSecurityRule
                    .ExternalRelationships, exception.Rule);
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void PresentationFacade_EnforcesPackageSecurityOnLegacyVba() {
            byte[] binary;
            using (PowerPointPresentation source =
                   PowerPointPresentation.Create()) {
                source.AddSlide().AddTextBox("Legacy VBA security policy");
                SetVbaProject(source, CreateVbaTestProject(
                    "SecurityModule", "Sub Main(): End Sub"));
                binary = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            var loadOptions = new PowerPointLoadOptions {
                PackageSecurity = OfficePackageSecurityOptions
                    .UntrustedDefaults
            };

            using var input = new MemoryStream(binary, writable: false);
            OfficePackageSecurityException exception = Assert.Throws<
                OfficePackageSecurityException>(() =>
                    PowerPointPresentation.Load(input, loadOptions));

            Assert.Equal(OfficePackageSecurityRule.Macros, exception.Rule);
        }

        [Fact]
        public void LegacyVbaConversion_EnforcesPackageSecurityBeforeOpeningPackage() {
            byte[] packageBytes;
            using (var package = new MemoryStream()) {
                using (var archive = new ZipArchive(package,
                           ZipArchiveMode.Create, leaveOpen: true)) {
                    ZipArchiveEntry vba = archive.CreateEntry(
                        "ppt/vbaProject.bin");
                    using Stream payload = vba.Open();
                    payload.WriteByte(1);
                }
                packageBytes = package.ToArray();
            }
            var loadOptions = new PowerPointLoadOptions {
                PackageSecurity = OfficePackageSecurityOptions
                    .UntrustedDefaults
            };

            OfficePackageSecurityException exception = Assert.Throws<
                OfficePackageSecurityException>(() =>
                    PowerPointPresentation
                        .ConvertProjectedVbaPackageToMacroEnabled(
                            packageBytes, loadOptions));

            Assert.Equal(OfficePackageSecurityRule.Macros, exception.Rule);
        }

        [Fact]
        public void PresentationFacade_EnforcesPackageSecurityOnOriginalLegacyContainer() {
            byte[] binary = CreatePresentationBytes();
            LegacyPptPresentation source = LegacyPptPresentation.Load(binary);
            byte[] withOpaqueObjectPool = source.Package
                .RewriteCompoundStreams(new Dictionary<string, byte[]> {
                    ["ObjectPool/Preserved/Contents"] =
                        new byte[] { 1, 2, 3, 4 }
                });
            var loadOptions = new PowerPointLoadOptions {
                PackageSecurity = OfficePackageSecurityOptions
                    .UntrustedDefaults
            };

            using var input = new MemoryStream(withOpaqueObjectPool,
                writable: false);
            OfficePackageSecurityException exception = Assert.Throws<
                OfficePackageSecurityException>(() =>
                    PowerPointPresentation.Load(input, loadOptions));

            Assert.Equal(OfficePackageSecurityRule.EmbeddedPayloads,
                exception.Rule);
        }

        [Theory]
        [InlineData(false, OfficePackageSecurityRule.EmbeddedPayloads)]
        [InlineData(true, OfficePackageSecurityRule.ActiveX)]
        public void PresentationFacade_EnforcesLegacyActiveContentPolicies(
            bool activeX, OfficePackageSecurityRule expectedRule) {
            byte[] storage = CreateOleTestStorage(activeX
                ? "ActiveX policy"
                : "OLE policy");
            byte[] binary;
            if (activeX) {
                binary = CreateExternalObjectFixture(storage,
                    ExternalObjectFixtureKind.ActiveX, compressed: false);
            } else {
                using PowerPointPresentation source =
                    PowerPointPresentation.Create();
                PowerPointSlide slide = source.AddSlide();
                using var payload = new MemoryStream(storage,
                    writable: false);
                slide.AddOleObject(payload, "Package");
                binary = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            var loadOptions = new PowerPointLoadOptions {
                PackageSecurity = OfficePackageSecurityOptions
                    .UntrustedDefaults
            };

            using var input = new MemoryStream(binary, writable: false);
            OfficePackageSecurityException exception = Assert.Throws<
                OfficePackageSecurityException>(() =>
                    PowerPointPresentation.Load(input, loadOptions));

            Assert.Equal(expectedRule, exception.Rule);
        }

        [Fact]
        public void PresentationFacade_RejectsPreserveOnlyLegacyExternalContent() {
            byte[] storage = CreateOleTestStorage(
                "Preserve-only linked OLE policy");
            byte[] binary = CreateExternalObjectFixture(storage,
                ExternalObjectFixtureKind.LinkedOle, compressed: false,
                linkedUpdateMode: uint.MaxValue);
            LegacyPptPresentation neutral = LegacyPptPresentation.Load(
                binary);
            Assert.Empty(neutral.LinkedOleObjects);
            Assert.Contains(neutral.Diagnostics, diagnostic =>
                diagnostic.Code.StartsWith("PPT-OLE-LINK-",
                    StringComparison.Ordinal));
            var security = OfficePackageSecurityOptions.SecureDefaults;
            security.ExternalRelationships =
                OfficePackageContentPolicy.Reject;
            var loadOptions = new PowerPointLoadOptions {
                PackageSecurity = security
            };

            using var input = new MemoryStream(binary, writable: false);
            OfficePackageSecurityException exception = Assert.Throws<
                OfficePackageSecurityException>(() =>
                    PowerPointPresentation.Load(input, loadOptions));

            Assert.Equal(OfficePackageSecurityRule.ExternalRelationships,
                exception.Rule);
        }

        [Fact]
        public void LegacyExternalPolicy_RecognizesLocationOnlyHyperlinks() {
            var hyperlink = new LegacyPptHyperlink(1, friendlyName: null,
                target: null, location: "https://example.test/location");

            Assert.True(PowerPointPresentation.IsExternalLegacyHyperlink(
                hyperlink));
        }

        [Fact]
        public void PresentationFacade_RejectsLegacyRunProgramActions() {
            var programUri = new Uri("file:///Applications/Calculator.app");
            byte[] binary;
            using (PowerPointPresentation source =
                   PowerPointPresentation.Create()) {
                PowerPointSlide slide = source.AddSlide(
                    P.SlideLayoutValues.Blank);
                PowerPointAutoShape shape = slide.AddRectangle(
                    100000, 100000, 1000000, 500000);
                HyperlinkRelationship relationship = slide.SlidePart
                    .AddHyperlinkRelationship(programUri, true);
                P.NonVisualDrawingProperties properties =
                    ((P.Shape)shape.Element).NonVisualShapeProperties!
                    .NonVisualDrawingProperties!;
                properties.Append(new A.HyperlinkOnClick {
                    Id = relationship.Id,
                    Action = "ppaction://program"
                });
                binary = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation neutral = LegacyPptPresentation.Load(
                binary);
            Assert.Contains(neutral.Slides[0].Shapes.SelectMany(shape =>
                    shape.Interactions), interaction =>
                interaction.Action ==
                    LegacyPptInteractionAction.RunProgram);
            var security = OfficePackageSecurityOptions.SecureDefaults;
            security.ExternalRelationships =
                OfficePackageContentPolicy.Reject;
            var loadOptions = new PowerPointLoadOptions {
                PackageSecurity = security
            };

            using var input = new MemoryStream(binary, writable: false);
            OfficePackageSecurityException exception = Assert.Throws<
                OfficePackageSecurityException>(() =>
                    PowerPointPresentation.Load(input, loadOptions));

            Assert.Equal(OfficePackageSecurityRule.ExternalRelationships,
                exception.Rule);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void PresentationFacade_RejectsPreserveOnlyLegacyRunProgramActions(
            bool corruptActionOwner) {
            byte[] binary;
            using (PowerPointPresentation source =
                   PowerPointPresentation.Create()) {
                PowerPointAutoShape shape = source.AddSlide(
                        P.SlideLayoutValues.Blank)
                    .AddRectangle(100000, 100000, 1000000, 500000);
                P.NonVisualDrawingProperties properties =
                    ((P.Shape)shape.Element).NonVisualShapeProperties!
                    .NonVisualDrawingProperties!;
                properties.Append(new A.HyperlinkOnClick {
                    Id = string.Empty,
                    Action = "ppaction://macro?name=Module1.RunReport"
                });
                binary = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            byte[] malformed = RewriteLegacyDocumentRecord(binary,
                record => record.Type == 0xF004
                    && record.DescendantsAndSelf().Any(child =>
                        child.Type == 0x0FF3
                        && child.PayloadLength == 16
                        && child.ReadByte(8) == (byte)
                            LegacyPptInteractionAction.Macro),
                (document, baseOffset, record) => {
                    LegacyPptRecord atom = record.DescendantsAndSelf()
                        .Single(child => child.Type == 0x0FF3);
                    document[checked(baseOffset + atom.PayloadOffset + 8)] =
                        (byte)LegacyPptInteractionAction.RunProgram;
                    LegacyPptRecord malformedRecord = corruptActionOwner
                        ? record.DescendantsAndSelf().Single(child =>
                            child.Type == 0xF011)
                        : record.Children.Single(child =>
                            child.Type is 0xF00F or 0xF010);
                    int malformedOffset = checked(baseOffset
                        + malformedRecord.Offset);
                    if (corruptActionOwner) {
                        WriteUInt16(document, malformedOffset,
                            unchecked((ushort)((malformedRecord.Instance
                                << 4) | 0x0E)));
                    } else {
                        WriteUInt16(document, malformedOffset + 2, 0xF012);
                    }
                });
            LegacyPptPresentation neutral = LegacyPptPresentation.Load(
                malformed);
            Assert.True(neutral.HasRunProgramContent);
            Assert.Empty(neutral.Slides[0].Shapes.SelectMany(shape =>
                shape.Interactions));
            using (var preservationInput = new MemoryStream(malformed,
                       writable: false))
            using (PowerPointPresentation imported =
                   PowerPointPresentation.Load(preservationInput,
                       new PowerPointLoadOptions {
                           LegacyPptImportOptions =
                               new LegacyPptImportOptions {
                                   ReportUnsupportedContent = false
                               }
                       })) {
                Assert.True(imported.LegacyPptWillPreserveRunProgramContent);
            }
            var security = OfficePackageSecurityOptions.SecureDefaults;
            security.ExternalRelationships =
                OfficePackageContentPolicy.Reject;

            using var input = new MemoryStream(malformed, writable: false);
            OfficePackageSecurityException exception = Assert.Throws<
                OfficePackageSecurityException>(() =>
                PowerPointPresentation.Load(input,
                    new PowerPointLoadOptions { PackageSecurity = security }));

            Assert.Equal(OfficePackageSecurityRule.ExternalRelationships,
                exception.Rule);
        }

        [Theory]
        [InlineData(0, false)]
        [InlineData(0, true)]
        [InlineData(1, false)]
        [InlineData(1, true)]
        [InlineData(2, false)]
        [InlineData(2, true)]
        [InlineData(3, false)]
        [InlineData(3, true)]
        public void LegacyPreservationGateScansProjectedShapeRoots(
            int targetKind, bool runProgram) {
            byte[] binary;
            using (PowerPointPresentation source =
                   PowerPointPresentation.Create()) {
                PowerPointSlide slide = source.AddSlide(
                    P.SlideLayoutValues.Blank);
                PowerPointAutoShape shape = slide.AddRectangle(
                    100000, 100000, 1000000, 500000);
                HyperlinkRelationship relationship = slide.SlidePart
                    .AddHyperlinkRelationship(new Uri(
                        "https://example.test/preserved"), true);
                ((P.Shape)shape.Element).NonVisualShapeProperties!
                    .NonVisualDrawingProperties!
                    .Append(new A.HyperlinkOnClick {
                        Id = relationship.Id,
                        Action = runProgram ? "ppaction://program" : null
                    });
                binary = source.ToBytes(PowerPointFileFormat.Ppt);
            }

            using var input = new MemoryStream(binary, writable: false);
            using PowerPointPresentation imported =
                PowerPointPresentation.Load(input);
            imported.Slides[0].SlidePart.Slide!
                .Descendants<A.HyperlinkOnClick>()
                .ToList()
                .ForEach(item => item.Remove());
            SlideLayoutPart layoutPart = imported.Slides[0].SlidePart
                .SlideLayoutPart!;
            DocumentFormat.OpenXml.OpenXmlPartRootElement root;
            OpenXmlPart targetPart;
            if (targetKind >= 2) {
                imported.Slides[0].Notes.Text = "Speaker note";
                NotesSlidePart notesPart = imported.Slides[0].SlidePart
                    .NotesSlidePart!;
                if (targetKind == 2) {
                    root = notesPart.NotesSlide!;
                    targetPart = notesPart;
                } else {
                    root = notesPart.NotesMasterPart!.NotesMaster!;
                    targetPart = notesPart.NotesMasterPart;
                }
            } else {
                if (targetKind == 0) {
                    root = layoutPart.SlideMasterPart!.SlideMaster!;
                    targetPart = layoutPart.SlideMasterPart;
                } else {
                    root = layoutPart.SlideLayout!;
                    targetPart = layoutPart;
                }
            }
            HyperlinkRelationship preservedRelationship = targetPart
                .AddHyperlinkRelationship(
                    new Uri("https://example.test/still-present"), true);
            P.NonVisualDrawingProperties properties = root
                .Descendants<P.NonVisualDrawingProperties>().First();
            properties.Append(new A.HyperlinkOnClick {
                Id = preservedRelationship.Id,
                Action = runProgram ? "ppaction://program" : null
            });

            Assert.Equal(runProgram,
                imported.LegacyPptWillPreserveRunProgramContent);
            Assert.Equal(!runProgram,
                imported.LegacyPptWillPreserveExternalHyperlinkContent);
        }

        [Fact]
        public void LegacyPreservationGateDoesNotTreatInternalSlideJumpAsExternal() {
            byte[] binary;
            using (PowerPointPresentation source =
                   PowerPointPresentation.Create()) {
                PowerPointSlide slide = source.AddSlide(
                    P.SlideLayoutValues.Blank);
                source.AddSlide(P.SlideLayoutValues.Blank);
                PowerPointAutoShape shape = slide.AddRectangle(
                    100000, 100000, 1000000, 500000);
                HyperlinkRelationship external = slide.SlidePart
                    .AddHyperlinkRelationship(
                        new Uri("https://example.test/remove"), true);
                ((P.Shape)shape.Element).NonVisualShapeProperties!
                    .NonVisualDrawingProperties!
                    .Append(new A.HyperlinkOnClick { Id = external.Id });
                binary = source.ToBytes(PowerPointFileFormat.Ppt);
            }

            using var input = new MemoryStream(binary, writable: false);
            using PowerPointPresentation imported =
                PowerPointPresentation.Load(input);
            SlidePart slidePart = imported.Slides[0].SlidePart;
            slidePart.Slide!.Descendants<A.HyperlinkOnClick>()
                .ToList().ForEach(item => item.Remove());
            HyperlinkRelationship internalJump = slidePart
                .AddHyperlinkRelationship(
                    new Uri("../slides/slide2.xml", UriKind.Relative),
                    false);
            P.NonVisualDrawingProperties properties = slidePart.Slide
                .Descendants<P.NonVisualDrawingProperties>().First();
            properties.Append(new A.HyperlinkOnClick {
                Id = internalJump.Id,
                Action = "ppaction://hlinksldjump"
            });

            Assert.False(imported.LegacyPptWillPreserveExternalHyperlinkContent);
        }

        [Fact]
        public void RawSecurityEvidenceScan_EnforcesRecordCountBudget() {
            var externalUri = new Uri("https://example.test/"
                + new string('a', 8192));
            byte[] binary;
            using (PowerPointPresentation source =
                   PowerPointPresentation.Create()) {
                PowerPointSlide slide = source.AddSlide(
                    P.SlideLayoutValues.Blank);
                PowerPointAutoShape shape = slide.AddRectangle(
                    100000, 100000, 1000000, 500000);
                HyperlinkRelationship relationship = slide.SlidePart
                    .AddHyperlinkRelationship(externalUri, true);
                ((P.Shape)shape.Element).NonVisualShapeProperties!
                    .NonVisualDrawingProperties!
                    .Append(new A.HyperlinkOnClick {
                        Id = relationship.Id
                    });
                binary = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            byte[] malformed = RewriteLegacyDocumentRecord(binary,
                record => record.Type == 0x0409,
                (document, baseOffset, record) => {
                    WriteUInt16(document,
                        checked(baseOffset + record.Offset),
                        unchecked((ushort)((record.Instance << 4) | 0x0E)));
                    Array.Clear(document,
                        checked(baseOffset + record.PayloadOffset),
                        record.PayloadLength);
                });

            InvalidDataException exception = Assert.Throws<
                InvalidDataException>(() => LegacyPptPresentation.Load(
                malformed, new LegacyPptImportOptions {
                    MaxRecordCount = 1000
                }));

            Assert.Contains("record count", exception.Message,
                StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void PresentationFacade_RejectsDuplicateLegacyExternalHyperlinkTargets() {
            var externalUri = new Uri("https://example.test/preserve-only");
            byte[] binary;
            using (PowerPointPresentation source =
                   PowerPointPresentation.Create()) {
                PowerPointSlide slide = source.AddSlide(
                    P.SlideLayoutValues.Blank);
                PowerPointAutoShape shape = slide.AddRectangle(
                    100000, 100000, 1000000, 500000);
                HyperlinkRelationship relationship = slide.SlidePart
                    .AddHyperlinkRelationship(externalUri, true);
                ((P.Shape)shape.Element).NonVisualShapeProperties!
                    .NonVisualDrawingProperties!
                    .Append(new A.HyperlinkOnClick {
                        Id = relationship.Id
                    });
                binary = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            byte[] malformed = RewriteLegacyDocumentRecord(binary,
                record => record.Type == 0x0FD7
                    && record.Children.Any(child => child.Type == 0x0FD3)
                    && record.Children.Any(child => child.Type == 0x0FBA
                        && child.Instance == 1),
                (document, baseOffset, record) => {
                    LegacyPptRecord atom = record.Children.Single(child =>
                        child.Type == 0x0FD3);
                    int atomOffset = checked(baseOffset + atom.Offset);
                    WriteUInt16(document, atomOffset, 0x0010);
                    WriteUInt16(document, atomOffset + 2, 0x0FBA);
                });
            LegacyPptPresentation neutral = LegacyPptPresentation.Load(
                malformed);
            Assert.True(neutral.HasExternalHyperlinkContent);
            Assert.Empty(neutral.Hyperlinks);
            Assert.Contains(neutral.Diagnostics, diagnostic =>
                diagnostic.Code == "PPT-HYPERLINK-ATOM");
            var security = OfficePackageSecurityOptions.SecureDefaults;
            security.ExternalRelationships =
                OfficePackageContentPolicy.Reject;

            using var input = new MemoryStream(malformed, writable: false);
            OfficePackageSecurityException exception = Assert.Throws<
                OfficePackageSecurityException>(() =>
                PowerPointPresentation.Load(input,
                    new PowerPointLoadOptions { PackageSecurity = security }));

            Assert.Equal(OfficePackageSecurityRule.ExternalRelationships,
                exception.Rule);
        }

        [Fact]
        public void EncryptedLegacyLoad_ValidatesOuterSourceBeforePasswordProcessing() {
            const string password = "source-policy-pass";
            byte[] encrypted;
            using (PowerPointPresentation source =
                   PowerPointPresentation.Create()) {
                source.AddSlide().AddTextBox("Encrypted source policy");
                encrypted = source.ToEncryptedBytes(password,
                    PowerPointFileFormat.Ppt);
            }
            OfficePackageSecurityReport report =
                OfficePackageSecurityInspector.Inspect(encrypted);
            Assert.Equal(OfficePackageContainerKind.CompoundBinary,
                report.ContainerKind);
            Assert.True(report.PartCount > 1);
            var security = OfficePackageSecurityOptions.SecureDefaults;
            security.MaxPartCount = report.PartCount - 1;
            var loadOptions = new PowerPointLoadOptions {
                PackageSecurity = security
            };

            using var input = new MemoryStream(encrypted, writable: false);
            OfficePackageSecurityException exception = Assert.Throws<
                OfficePackageSecurityException>(() =>
                    PowerPointPresentation.LoadEncrypted(input,
                        "wrong-password", loadOptions));

            Assert.Equal(OfficePackageSecurityRule.PartCount,
                exception.Rule);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void PresentationFacade_RejectsExternalHyperlinksWithMalformedOwners(
            bool corruptObjectList) {
            var externalUri = new Uri(
                "https://example.test/malformed-owner");
            byte[] binary;
            using (PowerPointPresentation source =
                   PowerPointPresentation.Create()) {
                PowerPointSlide slide = source.AddSlide(
                    P.SlideLayoutValues.Blank);
                PowerPointAutoShape shape = slide.AddRectangle(
                    100000, 100000, 1000000, 500000);
                HyperlinkRelationship relationship = slide.SlidePart
                    .AddHyperlinkRelationship(externalUri, true);
                ((P.Shape)shape.Element).NonVisualShapeProperties!
                    .NonVisualDrawingProperties!
                    .Append(new A.HyperlinkOnClick {
                        Id = relationship.Id
                    });
                binary = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            ushort ownerType = corruptObjectList
                ? (ushort)0x0409
                : (ushort)0x0FD7;
            byte[] malformed = RewriteLegacyDocumentRecord(binary,
                record => record.Type == ownerType,
                (document, baseOffset, record) => WriteUInt16(document,
                    checked(baseOffset + record.Offset),
                    unchecked((ushort)((record.Instance << 4) | 0x0E))));
            LegacyPptPresentation neutral = LegacyPptPresentation.Load(
                malformed);
            Assert.True(neutral.HasExternalHyperlinkContent);
            Assert.Empty(neutral.Hyperlinks);
            var security = OfficePackageSecurityOptions.SecureDefaults;
            security.ExternalRelationships =
                OfficePackageContentPolicy.Reject;

            using var input = new MemoryStream(malformed, writable: false);
            OfficePackageSecurityException exception = Assert.Throws<
                OfficePackageSecurityException>(() =>
                PowerPointPresentation.Load(input,
                    new PowerPointLoadOptions { PackageSecurity = security }));

            Assert.Equal(OfficePackageSecurityRule.ExternalRelationships,
                exception.Rule);
            Assert.Null(exception.ObservedValue);
        }

        [Fact]
        public void PresentationFacade_AllowsInternalHyperlinkInMalformedObjectList() {
            byte[] binary;
            using (PowerPointPresentation source =
                   PowerPointPresentation.Create()) {
                PowerPointSlide first = source.AddSlide(
                    P.SlideLayoutValues.Blank);
                PowerPointSlide destination = source.AddSlide(
                    P.SlideLayoutValues.Blank);
                PowerPointAutoShape shape = first.AddRectangle(
                    100000, 100000, 1000000, 500000);
                first.SlidePart.AddPart(destination.SlidePart);
                string relationshipId = first.SlidePart.GetIdOfPart(
                    destination.SlidePart);
                ((P.Shape)shape.Element).NonVisualShapeProperties!
                    .NonVisualDrawingProperties!
                    .Append(new A.HyperlinkOnClick {
                        Id = relationshipId,
                        Action = "ppaction://hlinksldjump"
                    });
                binary = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            byte[] malformed = RewriteLegacyDocumentRecord(binary,
                record => record.Type == 0x0409,
                (document, baseOffset, record) => WriteUInt16(document,
                    checked(baseOffset + record.Offset),
                    unchecked((ushort)((record.Instance << 4) | 0x0E))));
            LegacyPptPresentation neutral = LegacyPptPresentation.Load(
                malformed);
            Assert.False(neutral.HasExternalHyperlinkContent);
            var security = OfficePackageSecurityOptions.SecureDefaults;
            security.ExternalRelationships =
                OfficePackageContentPolicy.Reject;

            using var input = new MemoryStream(malformed, writable: false);
            using PowerPointPresentation loaded = PowerPointPresentation.Load(
                input, new PowerPointLoadOptions {
                    PackageSecurity = security
                });

            Assert.Equal(2, loaded.Slides.Count);
        }

        [Theory]
        [InlineData(0, OfficePackageSecurityRule.EmbeddedPayloads)]
        [InlineData(1, OfficePackageSecurityRule.ExternalRelationships)]
        [InlineData(2, OfficePackageSecurityRule.ActiveX)]
        [InlineData(3, OfficePackageSecurityRule.ExternalRelationships)]
        public void PresentationFacade_RejectsContentHiddenByMalformedExternalObjectList(
            int fixtureKind, OfficePackageSecurityRule expectedRule) {
            byte[] storage = CreateOleTestStorage(
                "Malformed external-object owner");
            byte[] binary;
            if (fixtureKind == 0) {
                using PowerPointPresentation source =
                    PowerPointPresentation.Create();
                using var payload = new MemoryStream(storage,
                    writable: false);
                source.AddSlide().AddOleObject(payload, "Package");
                binary = source.ToBytes(PowerPointFileFormat.Ppt);
            } else {
                ExternalObjectFixtureKind kind = fixtureKind switch {
                    1 => ExternalObjectFixtureKind.LinkedOle,
                    2 => ExternalObjectFixtureKind.ActiveX,
                    3 => ExternalObjectFixtureKind.LinkedWaveMedia,
                    _ => throw new ArgumentOutOfRangeException(
                        nameof(fixtureKind))
                };
                binary = CreateExternalObjectFixture(storage, kind,
                    compressed: false);
            }
            byte[] malformed = RewriteLegacyDocumentRecord(binary,
                record => record.Type == 0x0409,
                (document, baseOffset, record) => WriteUInt16(document,
                    checked(baseOffset + record.Offset),
                    unchecked((ushort)((record.Instance << 4) | 0x0E))));
            LegacyPptPresentation neutral = LegacyPptPresentation.Load(
                malformed);
            switch (fixtureKind) {
                case 0:
                    Assert.True(neutral.HasEmbeddedOleContent);
                    break;
                case 1:
                    Assert.True(neutral.HasLinkedOleContent);
                    break;
                case 2:
                    Assert.True(neutral.HasActiveXContent);
                    break;
                case 3:
                    Assert.True(neutral.HasExternalMediaContent);
                    break;
            }
            var security = OfficePackageSecurityOptions.SecureDefaults;
            switch (expectedRule) {
                case OfficePackageSecurityRule.EmbeddedPayloads:
                    security.EmbeddedPayloads =
                        OfficePackageContentPolicy.Reject;
                    break;
                case OfficePackageSecurityRule.ActiveX:
                    security.ActiveX = OfficePackageContentPolicy.Reject;
                    break;
                case OfficePackageSecurityRule.ExternalRelationships:
                    security.ExternalRelationships =
                        OfficePackageContentPolicy.Reject;
                    break;
                default:
                    throw new ArgumentOutOfRangeException(
                        nameof(expectedRule));
            }

            using var input = new MemoryStream(malformed, writable: false);
            OfficePackageSecurityException exception = Assert.Throws<
                OfficePackageSecurityException>(() =>
                PowerPointPresentation.Load(input,
                    new PowerPointLoadOptions { PackageSecurity = security }));

            Assert.Equal(expectedRule, exception.Rule);
            Assert.Null(exception.ObservedValue);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void PresentationFacade_RejectsVbaHiddenByMalformedOwnerOrDecoy(
            bool prependDecoy) {
            byte[] binary;
            using (PowerPointPresentation source =
                   PowerPointPresentation.Create()) {
                source.AddSlide().AddTextBox("Hidden VBA policy");
                SetVbaProject(source, CreateVbaTestProject(
                    "HiddenModule", "Sub Main(): End Sub"));
                binary = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            byte[] malformed;
            if (prependDecoy) {
                LegacyPptPresentation source = LegacyPptPresentation.Load(
                    binary);
                LegacyPptPersistObject documentPersist = source.Package
                    .PersistObjects[source.Package.DocumentPersistId];
                LegacyPptRecord document = LegacyPptRecordReader.ReadSingle(
                    documentPersist.RecordBytes, 0,
                    new LegacyPptImportOptions());
                var atomPayload = new byte[12];
                WriteVbaUInt32(atomPayload, 8, 2);
                byte[] decoy = BuildVbaRecord(version: 0x0F, instance: 0,
                    type: 0x07D0, payload: BuildVbaRecord(version: 0x0F,
                        instance: 1, type: 0x03FF,
                        payload: BuildVbaRecord(version: 2, instance: 0,
                            type: 0x0400, payload: atomPayload)));
                var children = new List<byte[]>();
                bool inserted = false;
                foreach (LegacyPptRecord child in document.Children) {
                    if (!inserted && child.Type == 0x07D0) {
                        children.Add(decoy);
                        inserted = true;
                    }
                    children.Add(child.CopyRecordBytes());
                }
                Assert.True(inserted);
                byte[] rewrittenDocument = BuildVbaRecord(
                    document.Version, document.Instance, document.Type,
                    JoinExternalObjectRecords(children));
                malformed = AppendLegacyDocumentPersist(source,
                    rewrittenDocument);
            } else {
                malformed = RewriteLegacyDocumentRecord(binary,
                    record => record.Type == 0x07D0,
                    (document, baseOffset, record) => WriteUInt16(document,
                        checked(baseOffset + record.Offset),
                        unchecked((ushort)((record.Instance << 4)
                            | 0x0E))));
            }
            LegacyPptPresentation neutral = LegacyPptPresentation.Load(
                malformed);
            Assert.True(neutral.HasVbaContent);
            Assert.Null(neutral.VbaProject);
            var security = OfficePackageSecurityOptions.SecureDefaults;
            security.Macros = OfficePackageContentPolicy.Reject;

            using var input = new MemoryStream(malformed, writable: false);
            OfficePackageSecurityException exception = Assert.Throws<
                OfficePackageSecurityException>(() =>
                PowerPointPresentation.Load(input,
                    new PowerPointLoadOptions { PackageSecurity = security }));

            Assert.Equal(OfficePackageSecurityRule.Macros, exception.Rule);
        }

        private static byte[] RewriteLegacyDocumentRecord(byte[] bytes,
            Func<LegacyPptRecord, bool> predicate,
            Action<byte[], int, LegacyPptRecord> rewrite) {
            LegacyPptPresentation source = LegacyPptPresentation.Load(bytes);
            byte[] document = (byte[])source.Package.DocumentStream.Clone();
            foreach (LegacyPptPersistObject persistObject in source.Package
                         .PersistObjects.Values) {
                LegacyPptRecord root = LegacyPptRecordReader.ReadSingle(
                    persistObject.RecordBytes, 0,
                    new LegacyPptImportOptions());
                LegacyPptRecord? record = root.DescendantsAndSelf()
                    .FirstOrDefault(predicate);
                if (record == null) continue;
                rewrite(document, checked((int)persistObject.StreamOffset),
                    record);
                return source.Package.RewriteCompoundStreams(
                    new Dictionary<string, byte[]>(
                        StringComparer.OrdinalIgnoreCase) {
                        ["PowerPoint Document"] = document
                    });
            }
            throw new InvalidDataException(
                "The requested legacy PowerPoint record was not found.");
        }

        private static byte[] AppendLegacyDocumentPersist(
            LegacyPptPresentation source, byte[] rewrittenDocument) {
            var offsets = source.Package.PersistObjectOffsets.ToDictionary(
                pair => pair.Key, pair => pair.Value);
            using var stream = new MemoryStream();
            stream.Write(source.Package.DocumentStream, 0,
                source.Package.DocumentStream.Length);
            offsets[source.Package.DocumentPersistId] = checked(
                (uint)stream.Position);
            stream.Write(rewrittenDocument, 0, rewrittenDocument.Length);
            uint directoryOffset = checked((uint)stream.Position);
            byte[] directory = BuildVbaPersistDirectory(offsets);
            stream.Write(directory, 0, directory.Length);
            uint editOffset = checked((uint)stream.Position);
            int oldEditOffset = checked((int)source.Package
                .CurrentEditOffset);
            int editLength = checked(8 + (int)ReadVbaUInt32(
                source.Package.DocumentStream, oldEditOffset + 4));
            var edit = new byte[editLength];
            Buffer.BlockCopy(source.Package.DocumentStream, oldEditOffset,
                edit, 0, edit.Length);
            WriteVbaUInt32(edit, 20, directoryOffset);
            stream.Write(edit, 0, edit.Length);

            byte[] currentUser = (byte[])source.Package.CurrentUserStream
                .Clone();
            WriteVbaUInt32(currentUser, 16, editOffset);
            return source.Package.RewriteCompoundStreams(
                new Dictionary<string, byte[]>(
                    StringComparer.OrdinalIgnoreCase) {
                    ["PowerPoint Document"] = stream.ToArray(),
                    ["Current User"] = currentUser
                });
        }

    }
}
