using System.IO;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core.Internal;
using OfficeIMO.Word;

namespace OfficeIMO.Security.Tests {
    public sealed class OfficeVbaWordBoundaryTests {
        [Theory]
        [InlineData(null)]
        [InlineData("1Normal.")]
        [InlineData("Unexpected.ThisDocument")]
        public void WordRejectsMissingOrForeignBaseBeforeAddingOrReplacingProject(string? identity) {
            OfficeVbaProject valid = OfficeVbaProject.Create();
            OfficeVbaModule module = valid.AddDocumentModuleWithIdentity("ThisDocument", "'original\r\n", "1Normal.ThisDocument");
            byte[] original = valid.Write().GetBytes();
            string source = module.Source.Replace("Attribute VB_Base = \"1Normal.ThisDocument\"\r\n",
                identity == null ? "" : "Attribute VB_Base = \"" + identity + "\"\r\n");
            OfficeVbaProject invalid = OfficeVbaProject.Load(ChangeStreams(original, new Dictionary<string, byte[]> {
                ["VBA/ThisDocument"] = OfficeVbaCompression.Compress(OfficeVbaText.Encode(source, 1252))
            }));
            foreach (bool existing in new[] { false, true }) {
                using WordDocument document = WordDocument.Create();
                if (existing) document.SetVbaProject(valid);
                byte[] before = document.ExtractMacros();
                Assert.Throws<ArgumentException>(() => document.SetVbaProject(invalid));
                Assert.Equal(before, document.ExtractMacros());
                Assert.Equal(existing, document.HasMacros);
            }
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void WordRemovesOpaqueModuleWhenDirectoryCompressionOrRecordsAreMalformed(bool records) {
            byte[] bytes = CreateModules();
            byte[] directory = records ? OfficeVbaCompression.Compress(new byte[] { 0xff }) : new byte[] { 2 };
            bytes = ChangeStreams(bytes, new Dictionary<string, byte[]> { ["VBA/dir"] = directory });
            using WordDocument document = WordDocument.Create();
            document.AddMacro(bytes);
            Assert.Equal(new[] { "Helpers", "Other" }, document.Macros.Select(module => module.Name));
            document.RemoveMacro("hElPeRs");
            Assert.Equal("Other", Assert.Single(document.Macros).Name);
            Assert.True(OfficeCompoundFileReader.TryRead(document.ExtractMacros(), out OfficeCompoundFile? compound, out _));
            Assert.Equal(directory, compound!.Streams["VBA/dir"]);
            Assert.False(compound.Streams.ContainsKey("VBA/Helpers"));
            document.RemoveMacro("Other");
            Assert.False(document.HasMacros);
        }

        [Fact]
        public void WordDoesNotUseOpaqueRemovalForUnreadableSourceWithValidDirectory() {
            byte[] bytes = ChangeStreams(CreateModules(), new Dictionary<string, byte[]> { ["VBA/Helpers"] = new byte[] { 2 } });
            using WordDocument document = WordDocument.Create();
            document.AddMacro(bytes);
            Assert.Throws<InvalidDataException>(() => document.RemoveMacro("Helpers"));
            Assert.Equal(bytes, document.ExtractMacros());
        }

        [Fact]
        public void WordOpaqueRemovalStillRequiresExplicitSignatureRemoval() {
            string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-word-opaque-" + Guid.NewGuid().ToString("N") + ".docm");
            byte[] bytes = ChangeStreams(CreateModules(), new Dictionary<string, byte[]> { ["VBA/dir"] = new byte[] { 2 } });
            try {
                using (WordDocument document = WordDocument.Create(path)) { document.AddMacro(bytes); document.Save(); }
                using (WordprocessingDocument package = WordprocessingDocument.Open(path, true)) {
                    ExtendedPart signature = package.MainDocumentPart!.VbaProjectPart!.AddExtendedPart(
                        "http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature",
                        "application/vnd.ms-office.vbaProjectSignature", ".bin");
                    using MemoryStream input = new MemoryStream(new byte[] { 1, 2, 3 });
                    signature.FeedData(input);
                }
                using WordDocument loaded = WordDocument.Load(path);
                Assert.Throws<InvalidOperationException>(() => loaded.RemoveMacro("Helpers"));
                Assert.Equal(bytes, loaded.ExtractMacros());
            } finally { if (File.Exists(path)) File.Delete(path); }
        }

        private static byte[] CreateModules() {
            OfficeVbaProject project = OfficeVbaProject.Create();
            project.AddModule("Helpers", "'first\r\n");
            project.AddModule("Other", "'second\r\n");
            return project.Write().GetBytes();
        }

        private static byte[] ChangeStreams(byte[] bytes, Dictionary<string, byte[]> replacements) {
            Assert.True(OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? compound, out _));
            return OfficeCompoundFileWriter.Rewrite(compound!, replacements);
        }
    }
}
