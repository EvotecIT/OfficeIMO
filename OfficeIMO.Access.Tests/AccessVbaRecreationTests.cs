using OfficeIMO.Access;
using System.Collections.Generic;

namespace OfficeIMO.Access.Tests {
    public sealed class AccessVbaRecreationTests {
        private static AccessDocument Load(string file) => AccessDocument.Load(Path.Combine(AppContext.BaseDirectory, "Fixtures", file));
        private static byte[] Save(AccessDocument document) { using MemoryStream output = new MemoryStream(); document.Save(output); return output.ToArray(); }
        private static string[] Permissions(AccessDocument document, int id) {
            var result = new List<string>();
            using AccessDataReader rows = document.SystemTables["MSysACEs"].OpenDataReader();
            while (rows.Read()) {
                if (Convert.ToInt32(rows.GetValue(rows.GetOrdinal("ObjectId"))) != id) continue;
                result.Add(Convert.ToBase64String((byte[])rows.GetValue(rows.GetOrdinal("SID"))) + ":"
                    + rows.GetValue(rows.GetOrdinal("ACM")) + ":" + rows.GetValue(rows.GetOrdinal("FInheritable")));
            }
            return result.OrderBy(x => x, StringComparer.Ordinal).ToArray();
        }

        [Theory]
        [InlineData("Application/objects-jet4.mdb", OfficeVbaModuleKind.Standard)]
        [InlineData("Application/objects-ace12.accdb", OfficeVbaModuleKind.Standard)]
        [InlineData("Application/objects-jet4.mdb", OfficeVbaModuleKind.Class)]
        [InlineData("Application/objects-ace12.accdb", OfficeVbaModuleKind.Class)]
        public void DeleteAndRecreateUsesFreshCatalogStorageAndPermissions(string file, OfficeVbaModuleKind kind) {
            using AccessDocument document = Load(file);
            AccessCatalogEntry original = Assert.Single(document.Catalog, x => x.NativeType == -32761);
            string[] permissions = Permissions(document, original.NativeId); Assert.NotEmpty(permissions);
            OfficeVbaProject stale = document.GetVbaProject();
            OfficeVbaProject project = document.GetVbaProject(); project.DeleteModule(original.Name);
            project.AddModule(original.Name, "Public Const Recreated As Long = 101\r\n", kind);
            document.SetVbaProject(project);
            AccessCatalogEntry replacement = Assert.Single(document.Catalog, x => x.NativeType == -32761);
            Assert.NotEqual(original.NativeId, replacement.NativeId); Assert.NotEqual(original.Id, replacement.Id);
            Assert.DoesNotContain(document.ApplicationStreams, x => x.Path == "Modules/0/PropData");
            Assert.Contains(document.ApplicationStreams, x => x.Path == "Modules/1/PropData");
            Assert.Empty(Permissions(document, original.NativeId)); Assert.Equal(permissions, Permissions(document, replacement.NativeId));
            project.SetModuleSource(original.Name, "Public Const Recreated As Long = 102\r\n"); document.SetVbaProject(project);
            Assert.Equal(replacement.NativeId, Assert.Single(document.Catalog, x => x.NativeType == -32761).NativeId);
            byte[] before = Save(document); long revision = document.Revision;
            stale.SetModuleSource(original.Name, "Public Const Stale As Long = 103\r\n");
            Assert.Throws<InvalidOperationException>(() => document.SetVbaProject(stale));
            Assert.Equal(revision, document.Revision); Assert.Equal(before, Save(document));
            using AccessDocument saved = AccessDocument.Load(new MemoryStream(before));
            Assert.Equal(replacement.NativeId, Assert.Single(saved.Catalog, x => x.NativeType == -32761).NativeId);
            Assert.Equal(kind, saved.GetVbaProject().GetModule(original.Name).Kind);
            Assert.Contains("Recreated As Long = 102", saved.GetVbaProject().GetModule(original.Name).Source);
            Assert.Empty(Permissions(saved, original.NativeId)); Assert.Equal(permissions, Permissions(saved, replacement.NativeId));
        }

        [Theory]
        [InlineData("Application/objects-jet4.mdb")]
        [InlineData("Application/objects-ace12.accdb")]
        public void IdenticalSourceRecreationStillChangesIdentityAndRollbackAllowsRetry(string file) {
            using AccessDocument document = Load(file);
            AccessCatalogEntry original = Assert.Single(document.Catalog, x => x.NativeType == -32761);
            OfficeVbaProject project = document.GetVbaProject(); OfficeVbaModule module = Assert.Single(project.Modules);
            string source = module.Source; project.DeleteModule(module.Name); project.AddModule(module.Name, source, module.Kind);
            byte[] before = Save(document);
            using (document.BeginUpdate()) {
                document.SetVbaProject(project);
                Assert.NotEqual(original.NativeId, Assert.Single(document.Catalog, x => x.NativeType == -32761).NativeId);
            }
            Assert.Equal(before, Save(document)); Assert.Same(original, Assert.Single(document.Catalog, x => x.NativeType == -32761));
            document.SetVbaProject(project);
            Assert.NotEqual(original.NativeId, Assert.Single(document.Catalog, x => x.NativeType == -32761).NativeId);
            using AccessDocument saved = AccessDocument.Load(new MemoryStream(Save(document)));
            Assert.Equal(source, saved.GetVbaProject().GetModule(module.Name).Source);
        }

        [Theory]
        [InlineData("Application/objects-jet4.mdb")]
        [InlineData("Application/objects-ace12.accdb")]
        public void RenameIntoDeletedNameRetainsOnlyTheSurvivingIdentity(string file) {
            using AccessDocument document = Load(file);
            AccessCatalogEntry original = Assert.Single(document.Catalog, x => x.NativeType == -32761);
            OfficeVbaProject project = document.GetVbaProject(); project.AddModule("Survivor", "Public Const SurvivorValue As Long = 104\r\n", OfficeVbaModuleKind.Class);
            document.SetVbaProject(project);
            AccessCatalogEntry survivor = document.Catalog.Single(x => x.Name == "Survivor");
            string[] permissions = Permissions(document, survivor.NativeId);
            project.DeleteModule(original.Name); project.RenameModule("Survivor", original.Name); document.SetVbaProject(project);
            AccessCatalogEntry final = Assert.Single(document.Catalog, x => x.NativeType == -32761);
            Assert.Equal(survivor.NativeId, final.NativeId); Assert.Equal(survivor.Id, final.Id);
            Assert.Empty(Permissions(document, original.NativeId)); Assert.Equal(permissions, Permissions(document, survivor.NativeId));
            Assert.DoesNotContain(document.ApplicationStreams, x => x.Path == "Modules/0/PropData");
            using AccessDocument saved = AccessDocument.Load(new MemoryStream(Save(document)));
            Assert.Equal(survivor.NativeId, Assert.Single(saved.Catalog, x => x.NativeType == -32761).NativeId);
            Assert.Equal(OfficeVbaModuleKind.Class, saved.GetVbaProject().GetModule(original.Name).Kind);
        }

        [Theory]
        [InlineData("Application/objects-jet4.mdb", false)]
        [InlineData("Application/objects-ace12.accdb", false)]
        [InlineData("Application/objects-jet4.mdb", true)]
        [InlineData("Application/objects-ace12.accdb", true)]
        public void AcceptedByteIdenticalProjectRemembersIdentityAcrossRenameOrReplacement(string file, bool replace) {
            using AccessDocument document = Load(file);
            AccessCatalogEntry original = Assert.Single(document.Catalog, x => x.NativeType == -32761);
            OfficeVbaProject accepted = OfficeVbaProject.Load(document.GetVbaProject().Write().GetBytes());
            byte[] initial = Save(document); long revision = document.Revision;
            using (document.BeginUpdate()) document.SetVbaProject(accepted);
            Assert.Equal(revision, document.Revision); Assert.Equal(initial, Save(document)); Assert.False(document.IsModified);
            OfficeVbaProject other = document.GetVbaProject();
            if (replace) { other.DeleteModule(original.Name); other.AddModule(original.Name, "Public Const Replacement As Long = 105\r\n"); }
            else other.RenameModule(original.Name, "InterveningName");
            document.SetVbaProject(other);
            byte[] before = Save(document); revision = document.Revision;
            accepted.SetModuleSource(original.Name, "Public Const Accepted As Long = 106\r\n");
            if (replace) {
                Assert.Throws<InvalidOperationException>(() => document.SetVbaProject(accepted));
                Assert.Equal(revision, document.Revision); Assert.Equal(before, Save(document));
            } else {
                document.SetVbaProject(accepted);
                AccessCatalogEntry final = Assert.Single(document.Catalog, x => x.NativeType == -32761);
                Assert.Equal(original.NativeId, final.NativeId); Assert.Equal(original.Id, final.Id);
                Assert.Contains(document.ApplicationStreams, x => x.Path == "Modules/0/PropData");
                using AccessDocument saved = AccessDocument.Load(new MemoryStream(Save(document)));
                Assert.Equal(original.NativeId, Assert.Single(saved.Catalog, x => x.NativeType == -32761).NativeId);
            }
        }

        [Theory]
        [InlineData("Application/objects-jet4.mdb")]
        [InlineData("Application/objects-ace12.accdb")]
        public void NoOpAcceptedDuringRolledBackAdditionCannotRecreateItsDiscardedIdentity(string file) {
            using AccessDocument document = Load(file); byte[] before = Save(document);
            OfficeVbaProject accepted;
            using (document.BeginUpdate()) {
                OfficeVbaProject added = document.GetVbaProject(); added.AddModule("RolledBack", "Public Const Value As Long = 107\r\n");
                document.SetVbaProject(added);
                accepted = OfficeVbaProject.Load(document.GetVbaProject().Write().GetBytes());
                long revision = document.Revision; AccessCatalogEntry[] staged = document.Catalog.ToArray();
                document.SetVbaProject(accepted);
                Assert.Equal(revision, document.Revision); Assert.Equal(staged, document.Catalog.ToArray());
            }
            Assert.Equal(before, Save(document)); long rolledBackRevision = document.Revision;
            accepted.SetModuleSource("RolledBack", "Public Const Value As Long = 108\r\n");
            Assert.Throws<InvalidOperationException>(() => document.SetVbaProject(accepted));
            Assert.Equal(rolledBackRevision, document.Revision); Assert.Equal(before, Save(document));
        }
    }
}
