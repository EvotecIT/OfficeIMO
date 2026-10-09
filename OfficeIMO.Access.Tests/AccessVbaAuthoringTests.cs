using OfficeIMO.Access;

namespace OfficeIMO.Access.Tests {
    public sealed class AccessVbaAuthoringTests {
        private static string Fixture(string name) => Path.Combine(AppContext.BaseDirectory, "Fixtures", name);
        private static byte[] Save(AccessDocument document) { using MemoryStream bytes = new MemoryStream(); document.Save(bytes); return bytes.ToArray(); }

        [Theory]
        [InlineData("Application/objects-jet4.mdb")]
        [InlineData("Application/objects-ace12.accdb")]
        [InlineData("Designer/designer-jet4.mdb")]
        [InlineData("Designer/designer-ace12.accdb")]
        public void NativeModuleAdditionPersistsSourceCatalogAndUnrelatedObjects(string fixture) {
            byte[] original = File.ReadAllBytes(Fixture(fixture));
            using AccessDocument document = AccessDocument.Load(new MemoryStream(original));
            OfficeVbaProject project = document.GetVbaProject();
            string oldName = project.Modules.Single().Name;
            AccessCatalogEntry originalModule = document.Catalog.Single(x => x.NativeType == -32761);
            project.SetModuleSource(oldName, project.GetModule(oldName).Source.Replace("Value = 42", "Value = 43"));
            project.AddModule("AddedModule", "Option Explicit\r\nPublic Function AddedValue() As Long\r\n AddedValue = 44\r\nEnd Function\r\n");
            project.AddModule("AddedClass", "Public Property Get Value() As Long\r\n Value = 45\r\nEnd Property\r\n", OfficeVbaModuleKind.Class);
            document.SetVbaProject(project);
            Assert.True(document.IsModified); Assert.Same(originalModule, document.Catalog.Single(x => x.Name == oldName && x.NativeType == -32761));
            Assert.Equal("vba.apply", Assert.Single(document.ChangeJournal).Operation);
            using AccessDocument saved = AccessDocument.Load(new MemoryStream(Save(document)));
            Assert.Equal(project.GetModule(oldName).Source.TrimEnd('\r','\n'), saved.VbaProject.Modules.Single(x => x.Name == oldName).Source!.TrimEnd('\r','\n'));
            Assert.Equal(project.GetModule("AddedModule").Source.TrimEnd('\r','\n'), saved.VbaProject.Modules.Single(x => x.Name == "AddedModule").Source!.TrimEnd('\r','\n'));
            Assert.Equal(originalModule.NativeId, saved.Catalog.Single(x => x.Name == oldName && x.NativeType == -32761).NativeId);
            Assert.Single(saved.Catalog, x => x.Name == "AddedModule" && x.NativeType == -32761);
            Assert.Equal(OfficeVbaModuleKind.Class, saved.GetVbaProject().GetModule("AddedClass").Kind);
            using (AccessDataReader rows = document.SystemTables.Single(x => x.Name == "MSysObjects").OpenDataReader()) {
                bool found = false; while (rows.Read()) if (Equals(rows["Name"], "AddedModule")) found = true;
                Assert.True(found);
            }
            Assert.Equal(document.Forms.Select(x => x.Name), saved.Forms.Select(x => x.Name));
            Assert.Equal(document.Reports.Select(x => x.Name), saved.Reports.Select(x => x.Name));
            using AccessDocument baseline = AccessDocument.Load(new MemoryStream(original));
            foreach (AccessStorageStream stream in baseline.ApplicationStreams.Where(x => !x.Path.StartsWith("VBA/", StringComparison.OrdinalIgnoreCase) && !x.Path.StartsWith("Modules/", StringComparison.OrdinalIgnoreCase)))
                Assert.Equal(stream.Payload.GetBytes(), saved.ApplicationStreams.Single(x => x.Path == stream.Path).Payload.GetBytes());
            Assert.Equal(original, File.ReadAllBytes(Fixture(fixture)));
        }

        [Theory]
        [InlineData("Application/objects-jet4.mdb")]
        [InlineData("Application/objects-ace12.accdb")]
        public void DetachedEditsAndRolledBackApplicationRetainExactSnapshotAndIdentities(string fixture) {
            byte[] original = File.ReadAllBytes(Fixture(fixture));
            using AccessDocument document = AccessDocument.Load(new MemoryStream(original));
            AccessCatalogEntry module = document.Catalog.Single(x => x.NativeType == -32761);
            AccessVbaProjectInfo inventory = document.VbaProject;
            OfficeVbaProject project = document.GetVbaProject();
            project.SetModuleSource(project.Modules.Single().Name, "Option Explicit\r\nPublic Const Replacement As Long = 7\r\n");
            Assert.False(document.IsModified); Assert.Equal(original, Save(document));
            using (document.BeginUpdate()) {
                document.SetVbaProject(project); Assert.True(document.HasActiveUpdate);
                Assert.Throws<InvalidOperationException>(() => document.AssessSave());
                Assert.Throws<InvalidOperationException>(() => document.Tables.First().OpenDataReader());
            }
            Assert.Same(inventory, document.VbaProject); Assert.Same(module, document.Catalog.Single(x => x.NativeType == -32761));
            Assert.False(document.IsModified); Assert.Empty(document.ChangeJournal); Assert.Equal(original, Save(document));
            document.SetVbaProject(document.GetVbaProject());
            Assert.False(document.IsModified); Assert.Equal(original, Save(document));
        }

        [Theory]
        [InlineData("Application/objects-jet4.mdb")]
        [InlineData("Application/objects-ace12.accdb")]
        public void NativeModuleDeletionRemovesCatalogAndSourceWhileKeepingUnrelatedPayloads(string fixture) {
            using AccessDocument document = AccessDocument.Load(Fixture(fixture));
            OfficeVbaProject project = document.GetVbaProject(); string name = project.Modules.Single().Name;
            project.DeleteModule(name); document.SetVbaProject(project);
            using AccessDocument saved = AccessDocument.Load(new MemoryStream(Save(document)));
            Assert.Empty(saved.VbaProject.Modules); Assert.DoesNotContain(saved.Catalog, x => x.NativeType == -32761);
            Assert.DoesNotContain(saved.ApplicationStreams, x => x.Path.EndsWith("/" + name, StringComparison.OrdinalIgnoreCase));
            Assert.Equal(document.Forms.Select(x => x.Name), saved.Forms.Select(x => x.Name));
        }

        [Fact]
        public void RejectedAndCancelledApplicationNeverChangesRevisionOrDestination() {
            using AccessDocument document = AccessDocument.Load(Fixture("Application/objects-ace12.accdb"));
            byte[] original = Save(document); OfficeVbaProject project = document.GetVbaProject();
            project.SetModuleSource(project.Modules.Single().Name, "Public Const Edited As Long = 9\r\n");
            using (AccessDataReader reader = document.SystemTables.First().OpenDataReader())
                Assert.Throws<InvalidOperationException>(() => document.SetVbaProject(project));
            using CancellationTokenSource cancelled = new CancellationTokenSource(); cancelled.Cancel();
            Assert.Throws<OperationCanceledException>(() => document.SetVbaProject(project, cancellationToken: cancelled.Token));
            Assert.Throws<InvalidDataException>(() => document.SetVbaProject(project, new OfficeVbaWriteOptions { MaximumProjectBytes = 1 }));
            Assert.False(document.IsModified); Assert.Equal(original, Save(document));
            document.SetVbaProject(project);
            using MemoryStream destination = new MemoryStream(); destination.WriteByte(42); destination.Position = 1;
            Assert.Throws<InvalidDataException>(() => document.Save(destination, new AccessSaveOptions { MaxOutputBytes = 4096 }));
            Assert.Equal(new byte[] { 42 }, destination.ToArray()); Assert.Equal(1, destination.Position);
        }

        [Fact]
        public void RepeatedNativeApplicationKeepsPriorEditsAndReadOnlyPolicy() {
            using AccessDocument document = AccessDocument.Load(Fixture("Application/objects-ace12.accdb"));
            OfficeVbaProject project = document.GetVbaProject(); project.AddModule("FirstAdded", "Public Const First As Long = 1\r\n");
            document.SetVbaProject(project); Guid firstId = document.Catalog.Single(x => x.Name == "FirstAdded").Id;
            OfficeVbaProject second = document.GetVbaProject(); second.AddModule("SecondAdded", "Public Const Second As Long = 2\r\n");
            document.SetVbaProject(second); Assert.Equal(firstId, document.Catalog.Single(x => x.Name == "FirstAdded").Id);
            using AccessDocument saved = AccessDocument.Load(new MemoryStream(Save(document)));
            Assert.Equal(3, saved.VbaProject.Modules.Count); Assert.Contains(saved.VbaProject.Modules, x => x.Name == "FirstAdded");
            Assert.Contains(saved.VbaProject.Modules, x => x.Name == "SecondAdded");
            using AccessDocument readOnly = AccessDocument.Load(new MemoryStream(Save(document)), new AccessLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
            Assert.Throws<InvalidOperationException>(() => readOnly.SetVbaProject(second));
        }

        [Theory]
        [InlineData("Application/objects-jet4.mdb")]
        [InlineData("Application/objects-ace12.accdb")]
        public void ModuleRenameRetainsNativeObjectIdentityAndItsDirectorySlot(string fixture) {
            using AccessDocument document = AccessDocument.Load(Fixture(fixture));
            OfficeVbaProject project = document.GetVbaProject(); string oldName = project.Modules.Single().Name;
            int oldId = document.Catalog.Single(x => x.NativeType == -32761).NativeId;
            Guid oldIdentity = document.Catalog.Single(x => x.NativeType == -32761).Id;
            project.RenameModule(oldName, "RenamedModule"); document.SetVbaProject(project);
            Assert.Equal(oldIdentity, document.Catalog.Single(x => x.NativeType == -32761).Id);
            using AccessDocument saved = AccessDocument.Load(new MemoryStream(Save(document)));
            Assert.Equal(oldId, saved.Catalog.Single(x => x.NativeType == -32761).NativeId);
            Assert.Equal("RenamedModule", Assert.Single(saved.VbaProject.Modules).Name);
            Assert.DoesNotContain(saved.ApplicationStreams, x => x.Path.EndsWith("/" + oldName, StringComparison.OrdinalIgnoreCase));
            Assert.Contains(saved.ApplicationStreams, x => x.Path == "Modules/0/PropData");
        }

        [Fact]
        public void BinaryReadFromEarlierAppliedSnapshotRemainsStableUntilDocumentDisposal() {
            using AccessDocument document = AccessDocument.Load(Fixture("Application/objects-ace12.accdb"));
            OfficeVbaProject project = document.GetVbaProject(); project.SetModuleSource("FixtureModule", "Public Const Value As Long = 1\r\n"); document.SetVbaProject(project);
            Stream? opened = null; byte[]? expected = null;
            using (AccessDataReader rows = document.SystemTables.Single(x => x.Name == "MSysAccessStorage").OpenDataReader()) while (rows.Read()) {
                if (!Equals(rows["Name"], "PROJECT")) continue;
                expected = (byte[])rows["Lv"]; opened = rows.GetStream(rows.GetOrdinal("Lv")); break;
            }
            Assert.NotNull(opened); Assert.NotNull(expected);
            using (Stream retained = opened!) {
                OfficeVbaProject second = document.GetVbaProject(); second.AddModule("AddedModule", "Public Const Value As Long = 2\r\n"); document.SetVbaProject(second);
                using MemoryStream read = new MemoryStream(); retained.CopyTo(read); Assert.Equal(expected, read.ToArray());
                document.Dispose(); Assert.Throws<ObjectDisposedException>(() => retained.ReadByte());
            }
        }

        [Fact]
        public void OwnAtomicSaveRefreshesSourceIdentityAndExternalChangesStillReject() {
            string root = Path.Combine(Path.GetTempPath(), "OfficeIMO-Access-Vba-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(root);
            try {
                string source = Path.Combine(root, "source.accdb"); File.Copy(Fixture("Application/objects-ace12.accdb"), source);
                using AccessDocument document = AccessDocument.Load(source);
                OfficeVbaProject project = document.GetVbaProject(); project.AddModule("AddedModule", "Public Const Value As Long = 4\r\n"); document.SetVbaProject(project);
                var replace = new AccessSaveOptions { FileConflictPolicy = OfficeConversionFileConflictPolicy.Replace };
                document.Save(source, replace); byte[] first = File.ReadAllBytes(source);
                document.ValidateSourceIdentity(); document.Save(source, replace); Assert.Equal(first, File.ReadAllBytes(source));
                File.Copy(Fixture("ace12.accdb"), source, true); byte[] external = File.ReadAllBytes(source);
                Assert.Throws<IOException>(() => document.Save(source, replace)); Assert.Equal(external, File.ReadAllBytes(source));
            } finally { Directory.Delete(root, true); }
        }
    }
}
