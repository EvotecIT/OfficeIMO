using OfficeIMO.Access;
using System.Collections.Generic;

namespace OfficeIMO.Access.Tests {
    public sealed class AccessApplicationTests {
        private static string Fixture(string name) => Path.Combine(AppContext.BaseDirectory, "Fixtures", name);

        [Theory]
        [InlineData("Application/objects-jet4.mdb")]
        [InlineData("Application/objects-ace12.accdb")]
        public void DeclaredVbaInventoryAndSourceAgreeWithAccessExports(string name) {
            using AccessDocument document = AccessDocument.Load(Fixture(name));
            Assert.Equal(AccessCatalogStatus.Decoded, document.VbaProject.CatalogStatus);
            AccessVbaModuleInfo module = Assert.Single(document.VbaProject.Modules);
            Assert.Equal("FixtureModule", module.Name); Assert.True(module.IsProcedural);
            Assert.True(module.SourceOffset > 0); Assert.Contains("VBA/VBAProject/VBA/", module.StoragePath);
            Assert.Equal("Attribute VB_Name = \"FixtureModule\"\r\n" + File.ReadAllText(Fixture(name + ".module.txt")).TrimEnd('\r','\n'), module.Source!.TrimEnd('\r','\n'));
            Assert.Contains(document.VbaProject.References, x => x.Name == "stdole");
            Assert.Contains(document.VbaProject.References, x => x.Name == "DAO");
            Assert.Contains(document.ApplicationStreams, x => x.Path.EndsWith("_VBA_PROJECT") && x.Payload.Length > 0);
            Assert.Contains(document.ApplicationStreams, x => x.Path == module.StoragePath);
            Assert.Equal(document.Catalog["FixtureForm"], document.Forms["FixtureForm"].CatalogEntry);
            Assert.Equal("Forms/0/", document.Forms["FixtureForm"].StoragePath);
            Assert.Single(document.Macros["FixtureMacro"].ActionMacro!.Actions, "StopMacro");
        }

        [Fact]
        public void ExpandedAceDesignerTreeAgreesWithIndependentNamedControlsAndSections() {
            using AccessDocument document = AccessDocument.Load(Fixture("Designer/designer-ace12.accdb"));
            AccessApplicationObject form = document.Forms["BoundForm1"]; AccessDesignerNode definition = Assert.IsType<AccessDesignerNode>(form.Definition);
            Assert.Equal("Synthetic form 1", definition.Caption); Assert.Equal("Contacts", definition.RecordSource); Assert.Equal(4800, definition.Width);
            AccessDesignerNode section = Assert.Single(definition.Children, x => x.NativeKind == 152);
            Assert.Equal("Detail", section.Name); Assert.Equal(3600, section.Height); Assert.Equal(3, section.Children.Count);
            Assert.Equal("Caption 1", section.Children.Single(x => x.Name == "Title1").Caption);
            AccessDesignerEvent click = Assert.Single(section.Children.Single(x => x.Name == "Title1").EventBindings);
            Assert.Equal("Click", click.Name); Assert.Equal("[Embedded Macro]", click.Expression);
            Assert.Single(click.EmbeddedMacro!.Actions, "StopMacro"); Assert.Equal(76, click.NativeEmbeddedMacro!.Length);
            Assert.Equal("DisplayName", section.Children.Single(x => x.Name == "DisplayName1").ControlSource);
            AccessDesignerNode combo = section.Children.Single(x => x.Name == "GroupChoice1"); Assert.Equal((ushort)111, combo.NativeKind);
            Assert.Equal("GroupId", combo.ControlSource); Assert.Equal("SELECT Id, Label FROM Groups; ", combo.RowSource);
            Assert.Equal("Forms/1/", document.Forms["BoundForm2"].StoragePath);
            Assert.Equal("Synthetic form 2", document.Forms["BoundForm2"].Definition!.Caption);
            Assert.Equal("Contacts", document.Reports["BoundReport"].Definition!.RecordSource);
            Assert.Contains(document.Dependencies, x => x.SourceObjectId == form.Id && x.Kind == "record-source" && x.TargetObjectId == document.Tables["Contacts"].Id);
            Assert.Contains(document.Dependencies, x => x.SourceObjectId == form.Id && x.Kind == "control-source" && x.TargetObjectId == document.Tables["Contacts"].Columns["DisplayName"].Id);
            Assert.Contains(document.Dependencies, x => x.Kind == "row-source" && x.TargetObjectId == null);
            Assert.Equal("Synthetic OfficeIMO application", document.Properties["AppTitle"]);
            Assert.Contains("Zażółć gęślą jaźń", document.VbaProject.Modules.Single().Source);
            Assert.Single(document.Macros["AutoExec"].ActionMacro!.Actions, "StopMacro");
            AccessDataMacroInfo dataMacro = Assert.Single(document.DataMacros);
            Assert.Equal(document.Catalog["Contacts"], dataMacro.CatalogEntry); Assert.Equal("AfterInsert", dataMacro.Event);
            Assert.Contains("Inert synthetic data macro", dataMacro.Xml); Assert.Single(dataMacro.Statements, "Comment");
            AccessResourceInfo resource = Assert.Single(document.Resources); Assert.Equal("Office Theme", resource.Name); Assert.Equal("thmx", resource.Type);
            Assert.NotNull(resource.Data); Assert.Single(resource.Data!.EnumerateAttachments());
        }

        [Theory]
        [InlineData("Application/objects-jet4.mdb", true)]
        [InlineData("Application/objects-ace12.accdb", true)]
        [InlineData("Designer/designer-jet4.mdb", true)]
        [InlineData("Designer/designer-ace12.accdb", true)]
        [InlineData("Profiles/password-jet4.mdb", true)]
        [InlineData("Profiles/password-ace.accdb", true)]
        [InlineData("Application/objects-ace12.accdb", false)]
        public void NoOpSavePreservesWholeFileAndEveryApplicationPayload(string name, bool decode) {
            byte[] original = File.ReadAllBytes(Fixture(name)); using MemoryStream input = new MemoryStream(original);
            using AccessDocument document = AccessDocument.Load(input, new AccessLoadOptions { DecodeCatalog = decode });
            Assert.Equal(AccessOperationStatus.Supported, document.AssessSave().Status);
            using MemoryStream output = new MemoryStream(); output.WriteByte(42); document.Save(output);
            Assert.Equal(original, output.ToArray()); Assert.Equal(0, output.Position); Assert.True(output.CanWrite);
            using AccessDocument reloaded = AccessDocument.Load(output, new AccessLoadOptions { DecodeCatalog = decode });
            Assert.Equal(document.Inspection!.Sha256, reloaded.Inspection!.Sha256);
            Assert.Equal(document.ApplicationStreams.Select(x => x.Path), reloaded.ApplicationStreams.Select(x => x.Path));
            foreach (AccessStorageStream stream in document.ApplicationStreams) Assert.Equal(stream.Payload.GetBytes(), reloaded.ApplicationStreams.Single(x => x.Path == stream.Path).Payload.GetBytes());
        }

        [Fact]
        public void NoOpPathSaveHonorsConflictsAndChangedSourceWithoutReplacingOutput() {
            string root = Path.Combine(Path.GetTempPath(), "OfficeIMO-Access-Preserve-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(root);
            try {
                string source = Path.Combine(root, "source.accdb"), target = Path.Combine(root, "target.accdb");
                byte[] original = File.ReadAllBytes(Fixture("Application/objects-ace12.accdb")); File.WriteAllBytes(source, original);
                using AccessDocument document = AccessDocument.Load(source); document.Save(target);
                Assert.Equal(original, File.ReadAllBytes(target));
                Assert.Throws<IOException>(() => document.Save(target)); Assert.Equal(original, File.ReadAllBytes(target));
                AccessSaveOptions options = new AccessSaveOptions { FileConflictPolicy = OfficeConversionFileConflictPolicy.Replace };
                document.Save(source, options); Assert.Equal(original, File.ReadAllBytes(source));
                File.WriteAllBytes(source, File.ReadAllBytes(Fixture("ace12.accdb")));
                Assert.Throws<IOException>(() => document.Save(target, options)); Assert.Equal(original, File.ReadAllBytes(target));
                using CancellationTokenSource canceled = new CancellationTokenSource(); canceled.Cancel();
                Assert.Throws<OperationCanceledException>(() => document.Save(Path.Combine(root,"canceled.accdb"), cancellationToken:canceled.Token));
                Assert.False(File.Exists(Path.Combine(root,"canceled.accdb")));
            } finally { Directory.Delete(root, true); }
        }

        [Fact]
        public void NativeApplicationMetadataHonorsLimitsAndImplicitDefaultsStayOpaque() {
            Assert.Throws<InvalidDataException>(() => AccessDocument.Load(Fixture("Application/objects-ace12.accdb"), new AccessLoadOptions { MaxMetadataBytes = 20_000 }));
            using AccessDocument document = AccessDocument.Load(Fixture("Application/objects-ace12.accdb"));
            Assert.IsType<AccessOpaqueValue>(document.Properties["ProjVer"]);
            Assert.Contains(document.Diagnostics, x => x.Code == "access.properties.opaque-value");
            AccessDesignerNode root = document.Forms["FixtureForm"].Definition!;
            Assert.Contains(root.Properties, x => x.Payload.Length == 0 && x.Value is AccessOpaqueValue);
            byte[] bytes = document.Forms["FixtureForm"].Streams.First(x => x.Path.EndsWith("Blob")).Payload.GetBytes(); bytes[0] = 0;
            Assert.Equal(21, document.Forms["FixtureForm"].Streams.First(x => x.Path.EndsWith("Blob")).Payload.GetBytes()[0]);
            Assert.Throws<NotSupportedException>(() => document.Tables.Add("CannotEditOpaqueSource"));
        }

        [Fact]
        public void RollbackRestoresChangeJournalAndStableObjectIdentity() {
            using AccessDocument document = AccessDocument.Create(); AccessTable table = document.Tables.Add("Journal"); table.Columns.Add("Id", AccessDataType.Int32);
            long revision = document.Revision; AccessChange[] before = document.ChangeJournal.ToArray();
            using (document.BeginUpdate()) { table.AppendRow(new AccessRowValues { ["Id"] = 1 }); Assert.Equal(table.Id, document.ChangeJournal.Last().ObjectId); Assert.Equal("row.append", document.ChangeJournal.Last().Operation); }
            Assert.Equal(revision, document.Revision); Assert.Equal(before, document.ChangeJournal); Assert.Equal(0, table.RowCount);
        }

        [Theory]
        [InlineData("Designer/designer-jet4.mdb")]
        [InlineData("Designer/designer-ace12.accdb")]
        public void InvalidApplicationDirectoryNameKeepsObjectsAndNativeBytesPreserveOnly(string name) {
            byte[] bytes = File.ReadAllBytes(Fixture(name));
            byte[] directory;
            using (AccessDocument original = AccessDocument.Load(new MemoryStream(bytes))) {
                directory = original.ApplicationStreams.Single(x => x.Path == "Forms/\u0003DirData").Payload.GetBytes();
                Assert.NotNull(original.Forms["BoundForm1"].StoragePath);
            }
            int match = FindUniquePayload(bytes, directory);
            Assert.Equal(4, directory[4]);
            // The first UTF-16 character follows the four-byte directory prefix and entry header.
            bytes[match + 6] = 0; bytes[match + 7] = 0xD8;
            directory[6] = 0; directory[7] = 0xD8;
            using MemoryStream input = new MemoryStream(bytes);
            using AccessDocument document = AccessDocument.Load(input);
            Assert.Equal(2, document.Forms.Count);
            Assert.All(document.Forms, form => {
                Assert.Null(form.StoragePath); Assert.Null(form.Definition); Assert.Empty(form.Streams);
                Assert.Contains(form.Diagnostics, x => x.Code == "access.application.preserve-only");
            });
            Assert.Equal(directory, document.ApplicationStreams.Single(x => x.Path == "Forms/\u0003DirData").Payload.GetBytes());
            Assert.NotNull(document.Reports["BoundReport"].StoragePath);
            Assert.Single(document.Macros["AutoExec"].ActionMacro!.Actions, "StopMacro");
            Assert.Contains("Zażółć", Assert.Single(document.VbaProject.Modules).Source);
            Assert.True(document.Tables["Contacts"].RowCount > 0);
            using MemoryStream output = new MemoryStream(); document.Save(output);
            Assert.Equal(bytes, output.ToArray());
        }

        [Theory]
        [InlineData("Designer/designer-jet4.mdb")]
        [InlineData("Designer/designer-ace12.accdb")]
        public void AliasedApplicationSlotsKeepTheirGroupPreserveOnly(string name) {
            byte[] bytes = File.ReadAllBytes(Fixture(name)); byte[] directory;
            using (AccessDocument original = AccessDocument.Load(new MemoryStream(bytes)))
                directory = original.ApplicationStreams.Single(x => x.Path == "Forms/\u0003DirData").Payload.GetBytes();
            int match = FindUniquePayload(bytes, directory), offset = 4;
            int firstSlot = BitConverter.ToInt32(directory, offset + 2 + directory[offset + 1] - 4);
            offset += 2 + directory[offset + 1]; int secondSlot = offset + 2 + directory[offset + 1] - 4;
            Assert.NotEqual(firstSlot, BitConverter.ToInt32(directory, secondSlot));
            Array.Copy(BitConverter.GetBytes(firstSlot), 0, bytes, match + secondSlot, 4);
            Array.Copy(BitConverter.GetBytes(firstSlot), 0, directory, secondSlot, 4);
            using MemoryStream input = new MemoryStream(bytes); using AccessDocument document = AccessDocument.Load(input);
            Assert.Equal(2, document.Forms.Count);
            Assert.All(document.Forms, form => { Assert.Null(form.StoragePath); Assert.Null(form.Definition); Assert.Empty(form.Streams); });
            Assert.Equal(directory, document.ApplicationStreams.Single(x => x.Path == "Forms/\u0003DirData").Payload.GetBytes());
            Assert.NotNull(document.Reports["BoundReport"].StoragePath); Assert.Single(document.Macros["AutoExec"].ActionMacro!.Actions, "StopMacro");
            Assert.Contains("Zażółć", Assert.Single(document.VbaProject.Modules).Source);
            using MemoryStream output = new MemoryStream(); document.Save(output); Assert.Equal(bytes, output.ToArray());
        }
        [Fact]
        public void DesignerParsingConsumesTheAggregateMetadataAllowance() {
            string file = Fixture("Designer/designer-ace12.accdb");
            using AccessDocument catalog = AccessDocument.Load(file, new AccessLoadOptions { MaxMetadataBytes = 44_000, DecodeApplicationObjects = false });
            Assert.True(catalog.Tables["Contacts"].RowCount > 0);
            InvalidDataException error = Assert.Throws<InvalidDataException>(() => AccessDocument.Load(file, new AccessLoadOptions { MaxMetadataBytes = 44_000 }));
            Assert.Contains("MaxMetadataBytes", error.Message);
        }
        [Fact]
        public void InvalidDesignerTextKeepsThePropertyOpaqueAndOtherDesignerMetadataReadable() {
            byte[] bytes = File.ReadAllBytes(Fixture("Designer/designer-ace12.accdb"));
            byte[] caption;
            using (AccessDocument original = AccessDocument.Load(new MemoryStream(bytes))) {
                caption = original.Forms["BoundForm1"].Definition!.Properties.Single(x => Equals(x.Value, "Synthetic form 1")).Payload.GetBytes();
            }
            int match = FindUniquePayload(bytes, caption);
            bytes[match] = 0; bytes[match + 1] = 0xD8;
            caption[0] = 0; caption[1] = 0xD8;
            using MemoryStream input = new MemoryStream(bytes);
            using AccessDocument document = AccessDocument.Load(input);
            AccessDesignerNode definition = Assert.IsType<AccessDesignerNode>(document.Forms["BoundForm1"].Definition);
            Assert.Null(definition.Caption); Assert.Equal("Contacts", definition.RecordSource);
            AccessDesignerProperty property = Assert.Single(definition.Properties, x => x.Payload.GetBytes().SequenceEqual(caption));
            Assert.IsType<AccessOpaqueValue>(property.Value);
            Assert.Equal("Synthetic form 2", document.Forms["BoundForm2"].Definition!.Caption);
            using MemoryStream output = new MemoryStream(); document.Save(output);
            Assert.Equal(bytes, output.ToArray());
        }

        [Fact]
        public void InvalidDataMacroTextKeepsTheDatabaseReadableAndNativePayloadExact() {
            byte[] bytes = File.ReadAllBytes(Fixture("Designer/designer-ace12.accdb"));
            byte[] marker = System.Text.Encoding.Unicode.GetBytes("Inert synthetic data macro");
            int match = FindUniquePayload(bytes, marker);
            bytes[match] = 0; bytes[match + 1] = 0xD8;
            byte[] macroPayload;
            using (AccessDocument original = AccessDocument.Load(new MemoryStream(File.ReadAllBytes(Fixture("Designer/designer-ace12.accdb"))))) {
                macroPayload = original.Catalog["Contacts"].NativePayloads["LvExtra"].GetBytes();
            }
            int payloadMatch = FindUniquePayload(macroPayload, marker);
            macroPayload[payloadMatch] = 0; macroPayload[payloadMatch + 1] = 0xD8;
            using MemoryStream input = new MemoryStream(bytes);
            using AccessDocument document = AccessDocument.Load(input);
            Assert.Empty(document.DataMacros);
            Assert.Contains(document.Diagnostics, x => x.Code == "access.data-macro.opaque");
            Assert.Equal(macroPayload, document.Catalog["Contacts"].NativePayloads["LvExtra"].GetBytes());
            Assert.True(document.Tables["Contacts"].RowCount > 0);
            Assert.Equal("Synthetic form 1", document.Forms["BoundForm1"].Definition!.Caption);
            Assert.Contains("Zażółć", Assert.Single(document.VbaProject.Modules).Source);
            using MemoryStream output = new MemoryStream(); document.Save(output);
            Assert.Equal(bytes, output.ToArray());
        }

        private static int FindUniquePayload(byte[] bytes, byte[] payload) {
            int match = -1;
            for (int offset = 0; offset <= bytes.Length - payload.Length; offset++) {
                int index = 0;
                while (index < payload.Length && bytes[offset + index] == payload[index]) index++;
                if (index != payload.Length) continue;
                Assert.Equal(-1, match); match = offset;
            }
            Assert.True(match >= 0); return match;
        }

        [Theory]
        [InlineData(0x100u)]
        [InlineData(uint.MaxValue)]
        public void UnknownDesignerPropertyRetainsItsFullNativeTypeAndPayload(uint nativeType) {
            byte[] bytes = File.ReadAllBytes(Fixture("Designer/designer-ace12.accdb"));
            byte[] caption = System.Text.Encoding.Unicode.GetBytes("Synthetic form 1");
            int match = FindUniquePayload(bytes, caption);
            // A designer property stores its four-byte type twelve bytes before its payload.
            int typeOffset = match - 12;
            for (int index = 0; index < 4; index++) bytes[typeOffset + index] = (byte)(nativeType >> (index * 8));
            using MemoryStream input = new MemoryStream(bytes);
            using AccessDocument document = AccessDocument.Load(input);
            AccessDesignerProperty property = Assert.Single(document.Forms["BoundForm1"].Definition!.Properties, x => x.NativeType == nativeType);
            Assert.Equal(nativeType, property.Payload.NativeType);
            AccessOpaqueValue value = Assert.IsType<AccessOpaqueValue>(property.Value);
            Assert.Equal(nativeType, value.NativeType); Assert.Equal(caption, value.GetBytes());
            Assert.Equal(caption, property.Payload.GetBytes());
            using MemoryStream output = new MemoryStream(); document.Save(output);
            Assert.Equal(bytes, output.ToArray());
        }
    }
}
