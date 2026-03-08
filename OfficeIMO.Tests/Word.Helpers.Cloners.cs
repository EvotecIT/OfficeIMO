using System.IO;
using System.Text;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_CopyFileStreamToMemoryStream_DoesNotKeepSourceLocked() {
            var sourcePath = Path.Combine(_directoryWithFiles, "ClonerSourceMemory.txt");
            File.WriteAllText(sourcePath, "memory");

            using var memoryStream = Helpers.CopyFileStreamToMemoryStream(sourcePath);

            File.Delete(sourcePath);

            Assert.False(File.Exists(sourcePath));
            Assert.Equal("memory", new StreamReader(memoryStream).ReadToEnd());
        }

        [Fact]
        public void Test_CopyFileStreamToFileStream_DoesNotKeepSourceLocked() {
            var sourcePath = Path.Combine(_directoryWithFiles, "ClonerSourceFile.txt");
            var destinationPath = Path.Combine(_directoryWithFiles, "ClonerDestinationFile.txt");
            File.WriteAllText(sourcePath, "file");

            using var destinationStream = Helpers.CopyFileStreamToFileStream(sourcePath, destinationPath);

            File.Delete(sourcePath);

            Assert.False(File.Exists(sourcePath));
            destinationStream.Position = 0;
            using var reader = new StreamReader(destinationStream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 1024, leaveOpen: true);
            Assert.Equal("file", reader.ReadToEnd());
        }
    }
}
