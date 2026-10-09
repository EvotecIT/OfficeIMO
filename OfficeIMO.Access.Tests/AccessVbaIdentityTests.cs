using OfficeIMO.Access;

namespace OfficeIMO.Access.Tests {
    public sealed class AccessVbaIdentityTests {
        private static AccessDocument Load(string name) => AccessDocument.Load(Path.Combine(AppContext.BaseDirectory, "Fixtures", name));
        private static byte[] Save(AccessDocument document) { using MemoryStream stream = new MemoryStream(); document.Save(stream); return stream.ToArray(); }

        [Theory]
        [InlineData("Application/objects-jet4.mdb")]
        [InlineData("Application/objects-ace12.accdb")]
        public void SourceEditAfterStagedRenameRetainsNativeStorageAndModelIdentity(string file) {
            using AccessDocument document = Load(file);
            AccessCatalogEntry original = document.Catalog.Single(x => x.NativeType == -32761);
            OfficeVbaProject project = document.GetVbaProject(); project.RenameModule(original.Name, "RenamedModule"); document.SetVbaProject(project);
            int? renamedStorageId = document.ApplicationStreams.Single(x => x.Path == "Modules/0/PropData").NativeId;
            project = document.GetVbaProject(); project.SetModuleSource("RenamedModule", "Public Const Marker As Long = 61\r\n"); document.SetVbaProject(project);
            AccessCatalogEntry renamed = document.Catalog.Single(x => x.NativeType == -32761);
            Assert.Equal(original.NativeId, renamed.NativeId); Assert.Equal(original.Id, renamed.Id);
            Assert.Equal(renamedStorageId, document.ApplicationStreams.Single(x => x.Path == "Modules/0/PropData").NativeId);
            using AccessDocument saved = AccessDocument.Load(new MemoryStream(Save(document)));
            Assert.Equal(original.NativeId, saved.Catalog.Single(x => x.NativeType == -32761).NativeId);
            Assert.Contains("Marker As Long = 61", saved.GetVbaProject().GetModule("RenamedModule").Source);
        }

        [Theory]
        [InlineData("Application/objects-jet4.mdb")]
        [InlineData("Application/objects-ace12.accdb")]
        public void RenamedIdentityAndNewModuleMayUseTheOriginalName(string file) {
            using AccessDocument document = Load(file);
            AccessCatalogEntry original = document.Catalog.Single(x => x.NativeType == -32761);
            OfficeVbaProject project = document.GetVbaProject(); project.RenameModule(original.Name, "RenamedModule");
            project.AddModule(original.Name, "Public Const NewIdentity As Long = 62\r\n"); document.SetVbaProject(project);
            Assert.Equal(2, document.Catalog.Count(x => x.NativeType == -32761));
            Assert.Equal(original.NativeId, document.Catalog.Single(x => x.Name == "RenamedModule").NativeId);
            Assert.NotEqual(original.NativeId, document.Catalog.Single(x => x.Name == original.Name).NativeId);
            using AccessDocument saved = AccessDocument.Load(new MemoryStream(Save(document)));
            Assert.Equal(2, saved.Catalog.Count(x => x.NativeType == -32761));
            Assert.Contains("NewIdentity As Long = 62", saved.GetVbaProject().GetModule(original.Name).Source);
        }

        [Theory]
        [InlineData("Application/objects-jet4.mdb")]
        [InlineData("Application/objects-ace12.accdb")]
        public void PendingAdditionRenameAndSiblingDeletionKeepTheSurvivingIdentity(string file) {
            using AccessDocument document = Load(file);
            OfficeVbaProject project = document.GetVbaProject(); project.AddModule("AddedFirst", "Public Const Value As Long = 1\r\n");
            project.AddModule("AddedSecond", "Public Const Value As Long = 2\r\n"); document.SetVbaProject(project);
            AccessCatalogEntry survivor = document.Catalog.Single(x => x.Name == "AddedSecond");
            project = document.GetVbaProject(); project.DeleteModule("AddedFirst"); project.RenameModule("AddedSecond", "SurvivingModule"); document.SetVbaProject(project);
            AccessCatalogEntry renamed = document.Catalog.Single(x => x.Name == "SurvivingModule");
            Assert.Equal(survivor.NativeId, renamed.NativeId); Assert.Equal(survivor.Id, renamed.Id);
            Assert.DoesNotContain(document.Catalog, x => x.Name == "AddedFirst");
            Assert.Equal(2, document.VbaProject.Modules.Count);
        }

        [Theory]
        [InlineData("Application/objects-jet4.mdb")]
        [InlineData("Application/objects-ace12.accdb")]
        public void ReusingDetachedProjectAfterApplicationKeepsItsPersistedIdentities(string file) {
            using AccessDocument document = Load(file);
            OfficeVbaProject project = document.GetVbaProject();
            string originalName = project.Modules.Single().Name;
            project.RenameModule(originalName, "RenamedModule");
            project.AddModule(originalName, "Public Const First As Long = 1\r\n");
            document.SetVbaProject(project);
            int originalId = document.Catalog.Single(x => x.Name == "RenamedModule").NativeId;
            int addedId = document.Catalog.Single(x => x.Name == originalName).NativeId;
            project.SetModuleSource("RenamedModule", "Public Const Second As Long = 2\r\n");
            document.SetVbaProject(project);
            Assert.Equal(originalId, document.Catalog.Single(x => x.Name == "RenamedModule").NativeId);
            project.RenameModule(originalName, "RenamedAddition");
            document.SetVbaProject(project);
            Assert.Equal(addedId, document.Catalog.Single(x => x.Name == "RenamedAddition").NativeId);
            using AccessDocument saved = AccessDocument.Load(new MemoryStream(Save(document)));
            Assert.Contains("Second As Long = 2", saved.GetVbaProject().GetModule("RenamedModule").Source);
        }

        [Theory]
        [InlineData("CodeBehind/empty-application-jet4.mdb")]
        [InlineData("CodeBehind/empty-application-ace12.accdb")]
        public void FirstProjectInExistingNativeApplicationRetainsBothModuleKinds(string file) {
            using AccessDocument document = Load(file);
            Assert.Empty(document.VbaProject.Modules);
            OfficeVbaProject project = OfficeVbaProject.Create("FirstProject");
            project.AddModule("FirstModule", "Public Const Marker As Long = 47\r\n");
            project.AddModule("FirstClass", "Public Property Get Marker() As Long\r\n Marker = 48\r\nEnd Property\r\n", OfficeVbaModuleKind.Class);
            document.SetVbaProject(project);
            using AccessDocument saved = AccessDocument.Load(new MemoryStream(Save(document)));
            Assert.Equal(2, saved.Catalog.Count(x => x.NativeType == -32761));
            Assert.Equal(OfficeVbaModuleKind.Class, saved.GetVbaProject().GetModule("FirstClass").Kind);
            Assert.Contains("Marker As Long = 47", saved.GetVbaProject().GetModule("FirstModule").Source);
        }

        [Theory]
        [InlineData("CodeBehind/code-behind-jet4.mdb")]
        [InlineData("CodeBehind/code-behind-ace12.accdb")]
        public void ExistingCodeBehindEditsPreserveHostBindingAndOrdinaryCatalog(string file) {
            using AccessDocument document = Load(file);
            string[] ordinary = document.Catalog.Where(x => x.NativeType == -32761).Select(x => x.Name).ToArray();
            OfficeVbaProject project = document.GetVbaProject();
            Assert.Equal(OfficeVbaModuleKind.Document, project.GetModule("Form_BoundForm1").Kind);
            Assert.Equal(OfficeVbaModuleKind.Document, project.GetModule("Report_BoundReport").Kind);
            project.SetModuleSource("Form_BoundForm1", project.GetModule("Form_BoundForm1").Source.Replace("marker 51", "marker 54"));
            project.SetModuleSource("Report_BoundReport", project.GetModule("Report_BoundReport").Source.Replace("marker 53", "marker 55"));
            document.SetVbaProject(project);
            Assert.Equal(ordinary, document.Catalog.Where(x => x.NativeType == -32761).Select(x => x.Name));
            using AccessDocument saved = AccessDocument.Load(new MemoryStream(Save(document)));
            Assert.Contains("marker 54", saved.GetVbaProject().GetModule("Form_BoundForm1").Source);
            Assert.Contains("marker 55", saved.GetVbaProject().GetModule("Report_BoundReport").Source);
            Assert.Equal(ordinary, saved.Catalog.Where(x => x.NativeType == -32761).Select(x => x.Name));
            foreach (AccessStorageStream stream in document.Forms["BoundForm1"].Streams.Concat(document.Reports["BoundReport"].Streams))
                Assert.Equal(stream.Payload.GetBytes(), saved.ApplicationStreams.Single(x => x.Path == stream.Path).Payload.GetBytes());
        }

        [Theory]
        [InlineData("CodeBehind/code-behind-jet4.mdb")]
        [InlineData("CodeBehind/code-behind-ace12.accdb")]
        public void HostStructuralEditsAndBaseRebindingFailWithoutDatabaseMutation(string file) {
            using AccessDocument document = Load(file); byte[] before = Save(document);
            OfficeVbaProject project = document.GetVbaProject();
            Assert.Throws<NotSupportedException>(() => project.RenameModule("Form_BoundForm1", "RenamedForm"));
            Assert.Throws<NotSupportedException>(() => project.DeleteModule("Report_BoundReport"));
            string original = project.GetModule("Form_BoundForm1").Source;
            string identity = original.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).Single(x => x.StartsWith("Attribute VB_Base = "));
            Assert.Throws<InvalidDataException>(() => project.SetModuleSource("Form_BoundForm1", original.Replace(identity, "Attribute VB_Base = \"0{00000000-0000-0000-0000-000000000000}\"")));
            Assert.Equal(original, project.GetModule("Form_BoundForm1").Source);
            Assert.Equal(before, Save(document)); Assert.False(document.IsModified);
        }
    }
}
