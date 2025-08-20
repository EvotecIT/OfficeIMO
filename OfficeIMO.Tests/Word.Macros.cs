using System;
using System.IO;
using System.Text;
using OpenMcdf;
using OfficeIMO.Word;
using Xunit;
using Version = OpenMcdf.Version;
using StorageModeFlags = OpenMcdf.StorageModeFlags;

namespace OfficeIMO.Tests {
    public partial class Word {
        private static string CreateDummyVba(string path, params string[] modules) {
            using var root = RootStorage.Create(path, Version.V3, StorageModeFlags.None);
            var vba = root.CreateStorage("VBA");
            using (var dir = vba.CreateStream("dir")) dir.Write(Array.Empty<byte>(), 0, 0);
            using (var project = vba.CreateStream("_VBA_PROJECT")) project.Write(Array.Empty<byte>(), 0, 0);
            if (modules == null || modules.Length == 0) modules = new[] { "Module1" };
            foreach (var module in modules) {
                using var stream = vba.CreateStream(module);
                var bytes = Encoding.UTF8.GetBytes("test");
                stream.Write(bytes, 0, bytes.Length);
            }
            using (var projectStream = root.CreateStream("PROJECT")) projectStream.Write(Array.Empty<byte>(), 0, 0);
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

        [Fact]
        public void Test_ExportImportMacrosBetweenDocuments() {
            string sourcePath = Path.Combine(_directoryWithFiles, "SourceMacroDoc.docm");
            string targetPath = Path.Combine(_directoryWithFiles, "TargetMacroDoc.docm");
            string macroFile = Path.Combine(_directoryWithFiles, "macroSource.bin");

            CreateDummyVba(macroFile);
            if (File.Exists(sourcePath)) File.Delete(sourcePath);
            if (File.Exists(targetPath)) File.Delete(targetPath);

            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddMacro(macroFile);
                document.Save();
            }

            byte[] macroData;
            using (WordDocument document = WordDocument.Load(sourcePath)) {
                Assert.True(document.HasMacros);
                macroData = document.ExtractMacros();
                Assert.NotNull(macroData);
                Assert.True(macroData.Length > 0);
            }

            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddMacro(macroData);
                Assert.True(document.HasMacros);
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(targetPath)) {
                Assert.True(document.HasMacros);
                Assert.Single(document.Macros);
                document.RemoveMacros();
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(targetPath)) {
                Assert.False(document.HasMacros);
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
            File.Delete(macroFile);
        }
    }
}
