using System.Collections.Generic;

namespace OfficeIMO.Access.Tests {
    public sealed class AccessVbaHostAuthoringTests {
        private static string Fixture(string name) => Path.Combine(AppContext.BaseDirectory, "Fixtures", "Designer", name);
        private static byte[] Save(AccessDocument document) { using MemoryStream bytes = new MemoryStream(); document.Save(bytes); return bytes.ToArray(); }
        private static IEnumerable<AccessDesignerNode> Nodes(AccessDesignerNode node) => new[] { node }.Concat(node.Children.SelectMany(Nodes));

        [Fact]
        public void EventCapabilityRequiresMutableDecodedAceApplication() {
            const string operation = "application.events.write";
            using AccessDocument modeled = AccessDocument.Create();
            using AccessDocument readOnly = AccessDocument.Load(Fixture("designer-ace12.accdb"), new AccessLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });
            using AccessDocument undecoded = AccessDocument.Load(Fixture("designer-ace12.accdb"), new AccessLoadOptions { DecodeApplicationObjects = false });
            using AccessDocument headerOnly = AccessDocument.Load(Fixture("designer-ace12.accdb"), new AccessLoadOptions { DecodeCatalog = false });
            using AccessDocument jet = AccessDocument.Load(Fixture("designer-jet4.mdb"));
            foreach (AccessDocument document in new[] { modeled, readOnly, undecoded, headerOnly, jet })
                Assert.False(document.Capabilities.Single(x => x.Operation == operation).IsSupported);
            using AccessDocument ace = AccessDocument.Load(Fixture("designer-ace12.accdb"));
            Assert.True(ace.Capabilities.Single(x => x.Operation == operation).IsSupported);
            ace.Forms["BoundForm1"].SetEventBinding(AccessEventKind.Open, "=Len(\"inert\")");
        }

        [Theory]
        [InlineData("DisplayName1", AccessEventKind.Click, 242U)]
        [InlineData("DisplayName1", AccessEventKind.AfterUpdate, 227U)]
        [InlineData("GroupChoice1", AccessEventKind.Click, 243U)]
        public void ControlEventsUseNativePerControlIdentityAndPreserveUnrelatedProperties(string name, AccessEventKind eventKind, uint id) {
            using AccessDocument document = AccessDocument.Load(Fixture("designer-ace12.accdb"));
            AccessApplicationObject host = document.Forms["BoundForm1"];
            AccessDesignerNode original = Nodes(host.Definition!).Single(x => x.Name == name);
            ushort code = eventKind == AccessEventKind.Click ? (ushort)126 : (ushort)86;
            var unchanged = original.Properties.Where(x => x.NativeCode != code).ToArray();
            host.SetEventBinding(eventKind, "=Len(\"inert-control\")", name);
            using AccessDocument saved = AccessDocument.Load(new MemoryStream(Save(document)));
            AccessDesignerNode node = Nodes(saved.Forms["BoundForm1"].Definition!).Single(x => x.Name == name);
            Assert.Equal(id, Assert.Single(node.Properties, x => x.NativeCode == code).NativeId);
            foreach (var property in unchanged) {
                AccessDesignerProperty retained = Assert.Single(node.Properties, x => x.NativeId == property.NativeId);
                Assert.Equal(property.Payload.GetBytes(), retained.Payload.GetBytes()); Assert.Equal(property.NativeCode, retained.NativeCode);
            }
            host.SetEventBinding(eventKind, "=Len(\"updated-control\")", name); host.SetEventBinding(eventKind, null, name);
            Assert.DoesNotContain(Nodes(host.Definition!).Single(x => x.Name == name).Properties, x => x.NativeCode == code);
        }

        [Fact]
        public void CodeBehindHonorsConfiguredReadAndWriteLimitsWithoutMutationOnRejection() {
            using AccessDocument document = AccessDocument.Load(Fixture("designer-ace12.accdb")); byte[] before = Save(document);
            AccessApplicationObject host = document.Forms["BoundForm1"];
            Assert.Throws<ArgumentOutOfRangeException>(() => host.SetCodeBehind("Public Const Inert As Long = 1\r\n", new OfficeVbaWriteOptions { MaximumExpandedBytes = 0 }));
            Assert.Throws<InvalidDataException>(() => host.SetCodeBehind("Public Const Inert As Long = 1\r\n", new OfficeVbaWriteOptions { MaximumExpandedBytes = 1 }));
            Assert.Throws<InvalidDataException>(() => host.SetCodeBehind("Public Const Inert As Long = 1\r\n", new OfficeVbaWriteOptions { MaximumProjectBytes = 1 }));
            Assert.Equal(before, Save(document)); Assert.False(document.IsModified);
            host.SetCodeBehind("Public Const Inert As Long = 2\r\n", new OfficeVbaWriteOptions { MaximumExpandedBytes = 128 * 1024, MaximumProjectBytes = 128 * 1024 });
            Assert.Contains("Inert As Long = 2", document.GetVbaProject().GetModule("Form_BoundForm1").Source);
        }

        [Theory]
        [InlineData("designer-jet4.mdb")]
        [InlineData("designer-ace12.accdb")]
        public void NewFormAndReportCodeBehindPreservesDesignerAndOrdinaryCatalog(string file) {
            using AccessDocument document = AccessDocument.Load(Fixture(file));
            var original = document.ApplicationStreams.ToDictionary(x => x.Path, x => x.Payload.GetBytes());
            int? moduleId = document.Catalog.Single(x => x.NativeType == -32761).NativeId;
            document.Forms["BoundForm1"].SetCodeBehind("Private Sub Form_Open(Cancel As Integer)\r\n 'inert marker 61\r\nEnd Sub\r\n");
            document.Reports["BoundReport"].SetCodeBehind("Private Sub Report_Open(Cancel As Integer)\r\n 'inert marker 62\r\nEnd Sub\r\n");
            OfficeVbaProject first = document.GetVbaProject(); string identity = first.GetModule("Form_BoundForm1").Source;
            document.Forms["BoundForm1"].SetCodeBehind("Public Const UpdatedSource As Long = 63\r\n");
            using AccessDocument saved = AccessDocument.Load(new MemoryStream(Save(document)));
            OfficeVbaProject project = saved.GetVbaProject();
            Assert.Equal(OfficeVbaModuleKind.Document, project.GetModule("Form_BoundForm1").Kind);
            Assert.Contains("UpdatedSource As Long = 63", project.GetModule("Form_BoundForm1").Source);
            Assert.Equal(identity.Split('\n').Single(x => x.Contains("VB_Base")), project.GetModule("Form_BoundForm1").Source.Split('\n').Single(x => x.Contains("VB_Base")));
            Assert.Contains("inert marker 62", project.GetModule("Report_BoundReport").Source);
            Assert.Equal(moduleId, Assert.Single(saved.Catalog, x => x.NativeType == -32761).NativeId);
            foreach (var stream in original.Where(x => !x.Key.StartsWith("VBA/") && x.Key != "Forms/0/PropData" && x.Key != "Reports/0/PropData"))
                Assert.Equal(stream.Value, saved.ApplicationStreams.Single(x => x.Path == stream.Key).Payload.GetBytes());
            Assert.Equal(1, document.Forms["BoundForm1"].Streams.Single(x => x.Path.EndsWith("PropData")).Payload.GetBytes()[9]);
            Assert.All(document.Forms["BoundForm1"].Diagnostics, x => Assert.Equal(document.Forms["BoundForm1"].Id, x.ObjectId));
        }

        [Theory]
        [InlineData("designer-jet4.mdb")]
        [InlineData("designer-ace12.accdb")]
        public void CodeBehindRollbackRestoresHostViewsAndExactDatabase(string file) {
            using AccessDocument document = AccessDocument.Load(Fixture(file)); byte[] before = Save(document);
            AccessApplicationObject host = document.Forms["BoundForm1"]; var streams = host.Streams; var definition = host.Definition;
            using (document.BeginUpdate()) host.SetCodeBehind("Public Const Rollback As Long = 64\r\n");
            Assert.Same(streams, host.Streams); Assert.Same(definition, host.Definition); Assert.Equal(before, Save(document));
            Assert.Empty(document.ChangeJournal); Assert.DoesNotContain(document.GetVbaProject().Modules, x => x.Name == "Form_BoundForm1");
            host.SetCodeBehind("Public Const Accepted As Long = 65\r\n");
            Assert.Contains("Accepted As Long = 65", document.GetVbaProject().GetModule("Form_BoundForm1").Source);
        }

        [Fact]
        public void EventInsertReplaceAndClearPreservesOtherDesignerPropertiesAndVbaBytes() {
            using AccessDocument document = AccessDocument.Load(Fixture("designer-ace12.accdb"));
            document.Forms["BoundForm1"].SetCodeBehind("Private Sub Form_Open(Cancel As Integer)\r\nEnd Sub\r\n");
            var vba = document.ApplicationStreams.Where(x => x.Path.StartsWith("VBA/")).ToDictionary(x => x.Path, x => x.Payload.GetBytes());
            AccessDesignerNode original = document.Forms["BoundForm1"].Definition!;
            document.Forms["BoundForm1"].SetEventBinding(AccessEventKind.Open, "[Event Procedure]");
            document.Forms["BoundForm1"].SetEventBinding(AccessEventKind.AfterUpdate, "=Len(\"inert\")", "GroupChoice1");
            document.Forms["BoundForm1"].SetEventBinding(AccessEventKind.Click, "=Len(\"replace\")", "Title1");
            byte[] set = Save(document); long revision = document.Revision;
            document.Forms["BoundForm1"].SetEventBinding(AccessEventKind.Click, "=Len(\"replace\")", "Title1");
            Assert.Equal(revision, document.Revision); Assert.Equal(set, Save(document));
            using AccessDocument saved = AccessDocument.Load(new MemoryStream(set));
            AccessDesignerNode updated = saved.Forms["BoundForm1"].Definition!;
            Assert.Equal("[Event Procedure]", Assert.Single(updated.EventBindings, x => x.Name == "Open").Expression);
            Assert.Equal("=Len(\"inert\")", Assert.Single(Nodes(updated).Single(x => x.Name == "GroupChoice1").EventBindings).Expression);
            foreach (AccessDesignerNode before in Nodes(original)) {
                AccessDesignerNode after = Nodes(updated).Single(x => x.Name == before.Name && x.NativeKind == before.NativeKind);
                foreach (AccessDesignerProperty property in before.Properties.Where(x => x.NativeCode != 77 && x.NativeCode != 86 && x.NativeCode != 126 && !(before.Name == "Title1" && x.NativeCode == 491)))
                    Assert.Equal(property.Payload.GetBytes(), after.Properties.Single(x => x.NativeId == property.NativeId).Payload.GetBytes());
            }
            foreach (var stream in vba) Assert.Equal(stream.Value, saved.ApplicationStreams.Single(x => x.Path == stream.Key).Payload.GetBytes());
            AccessDesignerNode label = Nodes(updated).Single(x => x.Name == "Title1");
            Assert.DoesNotContain(label.Properties, x => x.NativeCode == 491);
            Assert.Null(Assert.Single(label.EventBindings).NativeEmbeddedMacro);
            document.Forms["BoundForm1"].SetEventBinding(AccessEventKind.Open, null);
            using AccessDocument cleared = AccessDocument.Load(new MemoryStream(Save(document)));
            Assert.DoesNotContain(cleared.Forms["BoundForm1"].Definition!.EventBindings, x => x.Name == "Open");
        }

        [Fact]
        public void DesignerOnlyBindingDoesNotCreateVbaOrCodeBehindAndRollbackRestoresViews() {
            using AccessDocument document = AccessDocument.Load(Fixture("designer-ace12.accdb")); byte[] before = Save(document);
            AccessApplicationObject host = document.Forms["BoundForm2"]; var streams = host.Streams; var definition = host.Definition;
            using (document.BeginUpdate()) host.SetEventBinding(AccessEventKind.Click, "=Len(\"inert\")", "Title2");
            Assert.Equal(before, Save(document)); Assert.Same(streams, host.Streams); Assert.Same(definition, host.Definition);
            host.SetEventBinding(AccessEventKind.Click, "=Len(\"inert\")", "Title2");
            Assert.Single(document.GetVbaProject().Modules); Assert.Equal(0, host.Streams.Single(x => x.Path.EndsWith("PropData")).Payload.GetBytes()[9]);
        }

        [Fact]
        public void InvalidBindingsRejectWithoutChangingNativeBytes() {
            using AccessDocument document = AccessDocument.Load(Fixture("designer-ace12.accdb")); byte[] before = Save(document);
            AccessApplicationObject host = document.Forms["BoundForm1"];
            Assert.Throws<InvalidOperationException>(() => host.SetEventBinding(AccessEventKind.Open, "[Event Procedure]"));
            Assert.Throws<ArgumentException>(() => host.SetEventBinding(AccessEventKind.Click, "expression", "MissingControl"));
            Assert.Throws<NotSupportedException>(() => host.SetEventBinding(AccessEventKind.Click, "expression", "Detail"));
            Assert.Throws<NotSupportedException>(() => host.SetEventBinding(AccessEventKind.Click, "[Embedded Macro]", "Title1"));
            Assert.Throws<ArgumentException>(() => host.SetEventBinding(AccessEventKind.Open, "expression\0"));
            Assert.Throws<System.Text.EncoderFallbackException>(() => host.SetEventBinding(AccessEventKind.Open, "\ud800"));
            Assert.Throws<ArgumentOutOfRangeException>(() => host.SetEventBinding(AccessEventKind.Open, "expression", options: new OfficeVbaWriteOptions { MaximumProjectBytes = 0 }));
            Assert.Equal(before, Save(document)); Assert.False(document.IsModified);
            using AccessDocument jet = AccessDocument.Load(Fixture("designer-jet4.mdb")); byte[] jetBefore = Save(jet);
            Assert.Throws<NotSupportedException>(() => jet.Forms["BoundForm1"].SetEventBinding(AccessEventKind.Open, "expression"));
            Assert.Equal(jetBefore, Save(jet));
        }

        [Fact]
        public void EventOnlyApplicationCanReceiveItsFirstCodeBehindWithoutAnOrdinaryModule() {
            string path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "CodeBehind", "empty-designer-ace12.accdb");
            using AccessDocument document = AccessDocument.Load(path);
            Assert.Empty(document.VbaProject.Modules);
            var originalVba = document.ApplicationStreams.Where(x => x.Path.StartsWith("VBA/")).ToDictionary(x => x.Path, x => x.Payload.GetBytes());
            document.Forms["EventForm"].SetEventBinding(AccessEventKind.Click, "=Len(\"inert\")", "EventLabel");
            Assert.Empty(document.VbaProject.Modules);
            Assert.Equal(originalVba.Count, document.ApplicationStreams.Count(x => x.Path.StartsWith("VBA/")));
            foreach (var stream in originalVba) Assert.Equal(stream.Value, document.ApplicationStreams.Single(x => x.Path == stream.Key).Payload.GetBytes());
            document.Forms["EventForm"].SetCodeBehind("Private Sub EventLabel_Click()\r\n 'inert marker 71\r\nEnd Sub\r\n");
            document.Forms["EventForm"].SetEventBinding(AccessEventKind.Click, "[Event Procedure]", "EventLabel");
            using AccessDocument saved = AccessDocument.Load(new MemoryStream(Save(document)));
            Assert.Equal("Form_EventForm", Assert.Single(saved.GetVbaProject().Modules).Name);
            Assert.DoesNotContain(saved.Catalog, x => x.NativeType == -32761);
            Assert.Equal("[Event Procedure]", Assert.Single(Nodes(saved.Forms["EventForm"].Definition!).Single(x => x.Name == "EventLabel").EventBindings).Expression);
            saved.Forms["EventForm"].SetCodeBehind("Public Const ReopenedEdit As Long = 72\r\n");
            Assert.Contains("ReopenedEdit", saved.GetVbaProject().GetModule("Form_EventForm").Source);
        }
    }
}
