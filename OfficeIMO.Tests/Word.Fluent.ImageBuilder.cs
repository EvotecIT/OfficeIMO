using System.IO;
using System.Net;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO;
using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public async Task Test_FluentImageBuilderSources() {
            string filePath = Path.Combine(_directoryWithFiles, "FluentImageBuilderSources.docx");
            string imagePath = Path.Combine(_directoryWithImages, "Kulek.jpg");
            byte[] bytes = File.ReadAllBytes(imagePath);

            int port = GetAvailablePort();
            using var listener = new HttpListener();
            listener.Prefixes.Add($"http://localhost:{port}/");
            listener.Start();
            var serverTask = Task.Run(async () => {
                var context = await listener.GetContextAsync();
                context.Response.ContentType = "image/jpeg";
                context.Response.ContentLength64 = bytes.Length;
                await context.Response.OutputStream.WriteAsync(bytes, 0, bytes.Length);
                context.Response.OutputStream.Flush();
                context.Response.OutputStream.Close();
            });

            using (var document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .Image(i => i.Add(imagePath).Size(50, 50).Wrap(WrapTextImage.Square).Align(HorizontalAlignment.Center))
                    .Image(i => {
                        using var stream = File.OpenRead(imagePath);
                        i.Add(stream, "stream.jpg").Size(60, 60);
                    })
                    .Image(i => i.Add(bytes, "bytes.jpg").Size(70, 70))
                    .Image(i => i.AddFromUrl($"http://localhost:{port}/").Size(80, 80))
                    .End();
                document.Save(false);
            }

            await serverTask;
            listener.Stop();

            using (var document = WordDocument.Load(filePath)) {
                Assert.Equal(4, document.Images.Count);
                Assert.Equal(50, document.Images[0].Width);
                Assert.Equal(WrapTextImage.Square, document.Images[0].WrapText);
                Assert.Equal(JustificationValues.Center, document.Paragraphs[0].ParagraphAlignment);
                Assert.Equal(60, document.Images[1].Width);
                Assert.Equal(70, document.Images[2].Width);
                Assert.Equal(80, document.Images[3].Width);
            }
        }
    }
}
