using System;
using System.IO;
using System.IO.Compression;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioSaveApiTests {
        [Fact]
        public async Task SaveToBytesStreamAndPathUseTheSameVsdxContract() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            try {
                VisioDocument document = VisioDocument.Create(path);
                document.AddPage("Page-1").Shapes.Add(new VisioShape("1", 1, 1, 2, 1, "Start"));

                byte[] bytes = document.ToBytes();
                Assert.NotEmpty(bytes);
                using (var package = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read)) {
                    Assert.NotNull(package.GetEntry("visio/document.xml"));
                }

                using var stream = new MemoryStream();
                await document.SaveAsync(stream);
                await document.SaveAsync(path);
                Assert.Equal(bytes.Length, stream.Length);
                Assert.True(File.Exists(path));
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public async Task CanceledSaveDoesNotCreateOrReassociateDestination() {
            string originalPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string canceledPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            try {
                VisioDocument document = VisioDocument.Create(originalPath);
                document.AddPage("Page-1");
                using var cancellation = new CancellationTokenSource();
                cancellation.Cancel();

                await Assert.ThrowsAsync<OperationCanceledException>(() =>
                    document.SaveAsync(canceledPath, cancellation.Token));
                Assert.False(File.Exists(canceledPath));

                document.Save();
                Assert.True(File.Exists(originalPath));
            } finally {
                if (File.Exists(originalPath)) File.Delete(originalPath);
                if (File.Exists(canceledPath)) File.Delete(canceledPath);
            }
        }
    }
}
