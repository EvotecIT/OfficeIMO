using System.IO;
using System.Text;
using OpenMcdf;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        private static string CreateDummyVba(string path, params string[] modules) {
            var cf = new CompoundFile();
            var vba = cf.RootStorage.AddStorage("VBA");
            vba.AddStream("dir").SetData(new byte[0]);
            vba.AddStream("_VBA_PROJECT").SetData(new byte[0]);
            if (modules == null || modules.Length == 0) modules = new[] { "Module1" };
            foreach (var module in modules) {
                vba.AddStream(module).SetData(Encoding.UTF8.GetBytes("test"));
            }
            cf.RootStorage.AddStream("PROJECT").SetData(new byte[0]);
            cf.SaveAs(path);
            cf.Close();
            return path;
        }
        [Fact]
        public void Test_CreatingDocmWithMacro() {
            string macroPath = Path.Combine(_directoryDocuments, "vbaProject.bin");
            string filePath = Path.Combine(_directoryWithFiles, "DocumentWithMacro.docm");
            if (File.Exists(filePath)) File.Delete(filePath);

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddMacro(macroPath);
                Assert.True(document.HasMacros);
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.True(document.HasMacros);
                byte[] data = document.ExtractMacros();
                Assert.NotNull(data);
                Assert.True(data.Length > 0);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void Test_SavingAndRemovingMacros() {
            string macroPath = Path.Combine(_directoryDocuments, "vbaProject.bin");
            string filePath = Path.Combine(_directoryWithFiles, "DocumentWithMacro2.docm");
            if (File.Exists(filePath)) File.Delete(filePath);

            using (WordDocument document = WordDocument.Create(filePath)) {
                byte[] data = File.ReadAllBytes(macroPath);
                document.AddMacro(data);
                document.Save();
            }

            string extracted = Path.Combine(_directoryWithFiles, "macroCopy.bin");
            using (WordDocument document = WordDocument.Load(filePath)) {
                document.SaveMacros(extracted);
                Assert.True(File.Exists(extracted));
                document.RemoveMacros();
                Assert.False(document.HasMacros);
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.False(document.HasMacros);
            }

            File.Delete(filePath);
            File.Delete(extracted);
        }

        [Fact]
        public void Test_ListAndRemoveSingleMacro() {
            string vbaPath = Path.Combine(_directoryDocuments, "dummyVba.bin");
            CreateDummyVba(vbaPath);
            string filePath = Path.Combine(_directoryWithFiles, "DocumentWithMacro3.docm");
            if (File.Exists(filePath)) File.Delete(filePath);

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddMacro(vbaPath);
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Single(document.Macros);
                document.Macros[0].Remove();
                Assert.False(document.HasMacros);
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.False(document.HasMacros);
            }

            File.Delete(filePath);
            File.Delete(vbaPath);
        }

        [Fact]
        public void Test_RemoveOneOfMultipleMacros() {
            string vbaPath = Path.Combine(_directoryDocuments, "multiVba.bin");
            CreateDummyVba(vbaPath, "Module1", "Module2");
            string filePath = Path.Combine(_directoryWithFiles, "DocumentWithMacroMulti.docm");
            if (File.Exists(filePath)) File.Delete(filePath);

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddMacro(vbaPath);
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Equal(2, document.Macros.Count);
                document.RemoveMacro("Module1");
                Assert.True(document.HasMacros);
                Assert.Single(document.Macros);
                Assert.Equal("Module2", document.Macros[0].Name);
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.True(document.HasMacros);
                Assert.Single(document.Macros);
                Assert.Equal("Module2", document.Macros[0].Name);
            }

            File.Delete(filePath);
            File.Delete(vbaPath);
        }

        [Fact]
        public void Test_EnumeratingMacroNames() {
            string vbaPath = Path.Combine(_directoryDocuments, "dummyVba.bin");
            CreateDummyVba(vbaPath);
            using (WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "MacroNames.docm"))) {
                document.AddMacro(vbaPath);
                Assert.Single(document.Macros);
                Assert.Equal("Module1", document.Macros[0].Name);
            }
            File.Delete(vbaPath);
            File.Delete(Path.Combine(_directoryWithFiles, "MacroNames.docm"));
        }
    }
}
