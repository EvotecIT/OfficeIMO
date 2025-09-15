using System;
using System.IO;
using System.Net;
using System.Threading;
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
                        i.Add(stream, "stream.jpg").Size(60, 60).Align(HorizontalAlignment.Right);
                    })
                    .Image(i => i.Add(bytes, "bytes.jpg").Size(70, 70).Align(HorizontalAlignment.Left))
                    .Image(i => i.AddFromUrl($"http://localhost:{port}/").Size(80, 80).Align(HorizontalAlignment.Center))
                    .End();

                // Validate in-memory instead of reloading from disk
                Assert.Equal(4, document.Images.Count);
                Assert.Equal(50, document.Images[0].Width);
                Assert.Equal(WrapTextImage.Square, document.Images[0].WrapText);
                Assert.Equal(JustificationValues.Center, document.Paragraphs[0].ParagraphAlignment);
                Assert.Equal(60, document.Images[1].Width);
                Assert.Equal(JustificationValues.Right, document.Paragraphs[1].ParagraphAlignment);
                Assert.Equal(70, document.Images[2].Width);
                Assert.Equal(JustificationValues.Left, document.Paragraphs[2].ParagraphAlignment);
                Assert.Equal(80, document.Images[3].Width);
                Assert.Equal(JustificationValues.Center, document.Paragraphs[3].ParagraphAlignment);
            }

            await serverTask;
            listener.Stop();
        }

        [Fact]
        public async Task Test_FluentImageBuilderAddFromUrlCancellation() {
            string filePath = Path.Combine(_directoryWithFiles, "FluentImageBuilderCancelled.docx");

            using var document = WordDocument.Create(filePath);
            using var cts = new CancellationTokenSource();
            cts.Cancel();

            await Assert.ThrowsAnyAsync<OperationCanceledException>(async () => {
                await document.AsFluent().ImageAsync(async i => {
                    await i.AddFromUrlAsync("https://example.com/image.jpg", cts.Token);
                });
            });
        }

        [Fact]
        public async Task Test_FluentImageBuilderAddFromUrlInvalidScheme() {
            using var document = WordDocument.Create(Path.Combine(_directoryWithFiles, "FluentImageBuilderInvalid.docx"));

            await Assert.ThrowsAsync<ArgumentException>(async () => {
                await document.AsFluent().ImageAsync(async i => {
                    await i.AddFromUrlAsync("ftp://example.com/image.jpg");
                });
            });
        }

        [Fact]
        public async Task Test_FluentImageBuilderAddFromUrlNonImageContent() {
            string filePath = Path.Combine(_directoryWithFiles, "FluentImageBuilderNonImage.docx");

            int port = GetAvailablePort();
            using var listener = new HttpListener();
            listener.Prefixes.Add($"http://localhost:{port}/");
            listener.Start();
            var serverTask = Task.Run(async () => {
                var context = await listener.GetContextAsync();
                var data = System.Text.Encoding.UTF8.GetBytes("not image");
                context.Response.ContentType = "text/plain";
                context.Response.ContentLength64 = data.Length;
                await context.Response.OutputStream.WriteAsync(data, 0, data.Length);
                context.Response.OutputStream.Close();
            });

            using var document = WordDocument.Create(filePath);

            await Assert.ThrowsAsync<InvalidOperationException>(async () => {
                await document.AsFluent().ImageAsync(async i => {
                    await i.AddFromUrlAsync($"http://localhost:{port}/");
                });
            });

            await serverTask;
            listener.Stop();
        }

        [Fact]
        public void Test_FluentImageAltAndLink() {
            string filePath = Path.Combine(_directoryWithFiles, "FluentImageAltLink.docx");
            string imagePath = Path.Combine(_directoryWithImages, "Kulek.jpg");

            using (var document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .Image(i => i.Add(imagePath).Alt("Title", "Desc").Link("https://example.com"))
                    .End();
                var hyperlink = document.Paragraphs[0].Hyperlink!;
                var drawing = hyperlink._hyperlink.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>().First();
                var image = new WordImage(document, drawing);
                Assert.Equal("Desc", image.Description);
                Assert.Equal("Title", image.Title);
                Assert.True(document.Paragraphs[0].IsHyperLink);
            }
        }

        [Fact]
        public void Test_FluentImageCrop() {
            string filePath = Path.Combine(_directoryWithFiles, "FluentImageCrop.docx");
            string imagePath = Path.Combine(_directoryWithImages, "Kulek.jpg");

            using (var document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .Image(i => i.Add(imagePath).Crop(1, 2, 3, 4))
                    .End();
                Assert.Equal(1, document.Images[0].CropLeftCentimeters);
                Assert.Equal(2, document.Images[0].CropTopCentimeters);
                Assert.Equal(3, document.Images[0].CropRightCentimeters);
                Assert.Equal(4, document.Images[0].CropBottomCentimeters);
            }
        }

        [Fact]
        public void Test_FluentImageRotate() {
            string filePath = Path.Combine(_directoryWithFiles, "FluentImageRotate.docx");
            string imagePath = Path.Combine(_directoryWithImages, "Kulek.jpg");

            using (var document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .Image(i => i.Add(imagePath).Rotate(45))
                    .End();
                Assert.Equal(45, document.Images[0].Rotation);
            }
        }
    }
}
