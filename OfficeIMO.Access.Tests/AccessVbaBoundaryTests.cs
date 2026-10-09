using OfficeIMO.Access;

namespace OfficeIMO.Access.Tests {
    public sealed class AccessVbaBoundaryTests {
        private static string Fixture(string name) => Path.Combine(AppContext.BaseDirectory, "Fixtures", name);
        private static byte[] Save(AccessDocument document) { using MemoryStream output = new MemoryStream(); document.Save(output); return output.ToArray(); }

        [Theory]
        [InlineData("Application/objects-jet4.mdb")]
        [InlineData("Application/objects-ace12.accdb")]
        public void PermissionRowsUseTheirQualifiedPerObjectAllowance(string file) {
            using AccessDocument baseline = AccessDocument.Load(Fixture(file));
            int limit = Math.Max(baseline.Catalog.Count + 4, baseline.ApplicationStreams.Count * 2 + 16);
            Assert.True(baseline.SystemTables["MSysACEs"].RowCount > limit);
            using AccessDocument document = AccessDocument.Load(Fixture(file), new AccessLoadOptions { MaxCatalogObjects = limit });
            OfficeVbaProject project = document.GetVbaProject(); project.AddModule("WithinObjectLimit", "Public Const Inert As Long = 1\r\n");
            document.SetVbaProject(project);
            using AccessDocument saved = AccessDocument.Load(new MemoryStream(Save(document)), new AccessLoadOptions { MaxCatalogObjects = limit });
            Assert.Contains(saved.VbaProject.Modules, x => x.Name == "WithinObjectLimit");
            OfficeVbaProject removed = saved.GetVbaProject(); removed.DeleteModule("WithinObjectLimit"); saved.SetVbaProject(removed);
            Assert.DoesNotContain(saved.Catalog, x => x.Name == "WithinObjectLimit");
        }

        [Theory]
        [InlineData("Designer/designer-jet4.mdb", "Form_BoundForm1", OfficeVbaModuleKind.Standard)]
        [InlineData("Designer/designer-ace12.accdb", "Form_BoundForm1", OfficeVbaModuleKind.Standard)]
        [InlineData("Designer/designer-jet4.mdb", "Report_BoundReport", OfficeVbaModuleKind.Class)]
        [InlineData("Designer/designer-ace12.accdb", "Report_BoundReport", OfficeVbaModuleKind.Class)]
        public void OrdinaryHostLikeNamesRemainOrdinaryAndAllowUnrelatedEdits(string file, string name, OfficeVbaModuleKind kind) {
            using AccessDocument document = AccessDocument.Load(Fixture(file));
            OfficeVbaProject project = document.GetVbaProject(); project.AddModule(name, "Public Const Inert As Long = 2\r\n", kind);
            document.SetVbaProject(project);
            using AccessDocument saved = AccessDocument.Load(new MemoryStream(Save(document)));
            OfficeVbaProject next = saved.GetVbaProject(); next.SetModuleSource(name, "Public Const Inert As Long = 3\r\n"); saved.SetVbaProject(next);
            Assert.Equal(kind, saved.GetVbaProject().GetModule(name).Kind);
            Assert.Equal(name, Assert.Single(saved.Catalog, x => x.NativeType == -32761 && x.Name == name).Name);
            Assert.All(saved.Forms.Concat(saved.Reports), x => Assert.Equal(0, x.Streams.Single(s => s.Path.EndsWith("PropData")).Payload.GetBytes()[9]));
        }

        [Theory]
        [InlineData("Application/objects-jet4.mdb", "DigitalSignature")]
        [InlineData("Application/objects-ace12.accdb", "DigitalSignature")]
        [InlineData("Application/objects-jet4.mdb", "MyDigitalSignatureModule")]
        [InlineData("Application/objects-ace12.accdb", "MyDigitalSignatureModule")]
        public void SignatureLikeDeclaredModuleStreamsAcceptStagedAndReopenedEdits(string file, string name) {
            using AccessDocument document = AccessDocument.Load(Fixture(file));
            OfficeVbaProject project = document.GetVbaProject(); project.AddModule(name, "Public Const Inert As Long = 4\r\n"); document.SetVbaProject(project);
            project.SetModuleSource(name, "Public Const Inert As Long = 5\r\n"); document.SetVbaProject(project);
            using AccessDocument saved = AccessDocument.Load(new MemoryStream(Save(document)));
            OfficeVbaProject next = saved.GetVbaProject(); next.SetModuleSource(name, "Public Const Inert As Long = 6\r\n"); saved.SetVbaProject(next);
            Assert.Contains("Inert As Long = 6", saved.GetVbaProject().GetModule(name).Source);
        }

        [Theory]
        [InlineData("CodeBehind/empty-application-jet4.mdb")]
        [InlineData("CodeBehind/empty-application-ace12.accdb")]
        public void EmptyFirstProjectAllowsFirstModuleAndReadditionAfterLastRemoval(string file) {
            using AccessDocument document = AccessDocument.Load(Fixture(file));
            document.SetVbaProject(OfficeVbaProject.Create());
            using AccessDocument saved = AccessDocument.Load(new MemoryStream(Save(document)));
            OfficeVbaProject first = saved.GetVbaProject(); first.AddModule("FirstModule", "Public Const Inert As Long = 7\r\n"); saved.SetVbaProject(first);
            OfficeVbaProject removed = saved.GetVbaProject(); removed.DeleteModule("FirstModule"); saved.SetVbaProject(removed);
            OfficeVbaProject second = saved.GetVbaProject(); second.AddModule("SecondModule", "Public Const Inert As Long = 8\r\n"); saved.SetVbaProject(second);
            Assert.Equal("SecondModule", Assert.Single(saved.GetVbaProject().Modules).Name);
        }

        [Theory]
        [InlineData("Application/objects-jet4.mdb")]
        [InlineData("Application/objects-ace12.accdb")]
        public void RebuiltIndexesConsumeRemainingAggregateMetadataWithoutPoisoningRetry(string file) {
            byte[] input = File.ReadAllBytes(Fixture(file)); int low = 1, high = 4 * 1024 * 1024;
            while (low < high) {
                int midpoint = low + (high - low) / 2; bool decoded;
                try { using AccessDocument probe = AccessDocument.Load(new MemoryStream(input), new AccessLoadOptions { MaxMetadataBytes = midpoint }); decoded = probe.VbaProject.CatalogStatus == AccessCatalogStatus.Decoded; }
                catch (InvalidDataException) { decoded = false; }
                if (decoded) high = midpoint; else low = midpoint + 1;
            }
            using AccessDocument document = AccessDocument.Load(new MemoryStream(input), new AccessLoadOptions { MaxMetadataBytes = low + (file.EndsWith("accdb") ? 4096 : 1024) });
            OfficeVbaProject renamed = document.GetVbaProject(); string name = Assert.Single(renamed.Modules).Name;
            renamed.RenameModule(name, new string('X', name.Length));
            InvalidDataException error = Assert.Throws<InvalidDataException>(() => document.SetVbaProject(renamed));
            Assert.True(error.Message.Contains("aggregate metadata"), error.ToString()); Assert.Equal(input, Save(document)); Assert.False(document.IsModified);
            OfficeVbaProject source = document.GetVbaProject(); source.SetModuleSource(name, source.GetModule(name).Source.Replace("Value = 42", "Value = 43")); document.SetVbaProject(source);
            Assert.Contains("Value = 43", document.GetVbaProject().GetModule(name).Source);
        }
    }
}
