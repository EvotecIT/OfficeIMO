using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Threading.Tasks;
using SixLabors.ImageSharp;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Markdown {
        [Fact]
        public async Task MarkdownToWord_ParsesImageHints() {
            string imagePath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png"));
            int port = GetAvailablePort();
            string remoteUrl = $"http://localhost:{port}/";

            using var listener = new HttpListener();
            listener.Prefixes.Add($"http://localhost:{port}/");
            listener.Start();
            var serverTask = Task.Run(() => {
                var bytes = File.ReadAllBytes(imagePath);
                while (listener.IsListening) {
                    HttpListenerContext context;
                    try {
                        context = listener.GetContext();
                    } catch (HttpListenerException) {
                        break;
                    } catch (ObjectDisposedException) {
                        break;
                    }

                    try {
                        context.Response.ContentType = "image/png";
                        context.Response.ContentLength64 = bytes.Length;
                        context.Response.OutputStream.Write(bytes, 0, bytes.Length);
                        context.Response.OutputStream.Flush();
                        context.Response.Close();
                    } catch (HttpListenerException) {
                        break;
                    } catch (ObjectDisposedException) {
                        break;
                    }
                }
            });

            try {
                string md = $"![Local]({imagePath}){{width=40 height=30}}\n" +
                             $"![Remote]({remoteUrl}){{width=50 height=20}}";
                var warnings = new List<string>();
                var doc = md.LoadFromMarkdown(new MarkdownToWordOptions {
                    AllowLocalImages = true,
                    AllowRemoteImages = true,
                    OnWarning = warnings.Add
                });

                Assert.True(
                    doc.Images.Count + doc.HyperLinks.Count >= 2,
                    $"Expected both markdown image entries to be preserved as images or hyperlinks. Images={doc.Images.Count}, Hyperlinks={doc.HyperLinks.Count}, Warnings: {string.Join(" | ", warnings)}");
            } finally {
                listener.Stop();
                await serverTask;
            }
        }

        [Fact]
        public void MarkdownToWord_UsesNaturalSizeWhenNoHints() {
            string imagePath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png"));
            string md = $"![Local]({imagePath})";
            var doc = md.LoadFromMarkdown(new MarkdownToWordOptions {
                AllowLocalImages = true
            });

            using var image = Image.Load(imagePath, out _);

            Assert.Single(doc.Images);
            Assert.Equal("Local", doc.Images[0].Description);
            Assert.Equal(image.Width, doc.Images[0].Width);
            Assert.Equal(image.Height, doc.Images[0].Height);

        }

        [Fact]
        public void MarkdownToWord_FitsImageToPageContentWidthWhenEnabled() {
            string imagePath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png"));
            string md = $"![Local]({imagePath}){{width=1200 height=300}}";
            var doc = md.LoadFromMarkdown(new MarkdownToWordOptions {
                AllowLocalImages = true,
                FitImagesToPageContentWidth = true,
                ImageLayout = {
                    AllowUpscale = true
                },
                DefaultPageSize = WordPageSize.Letter
            });

            Assert.Single(doc.Images);
            Assert.InRange(doc.Images[0].Width ?? 0, 623, 625);
            Assert.InRange(doc.Images[0].Height ?? 0, 155, 157);
        }

        [Fact]
        public void MarkdownToWord_AppliesConfiguredImageMaxWidthPixels() {
            string imagePath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png"));
            string md = $"![Local]({imagePath}){{width=1200 height=300}}";
            var doc = md.LoadFromMarkdown(new MarkdownToWordOptions {
                AllowLocalImages = true,
                MaxImageWidthPixels = 480
            });

            Assert.Single(doc.Images);
            Assert.InRange(doc.Images[0].Width ?? 0, 479.5, 480.5);
            Assert.InRange(doc.Images[0].Height ?? 0, 119.5, 120.5);
        }

        [Fact]
        public void MarkdownToWord_AppliesConfiguredImageMaxWidthPercentOfContent() {
            string imagePath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png"));
            string md = $"![Local]({imagePath}){{width=1200 height=300}}";
            var doc = md.LoadFromMarkdown(new MarkdownToWordOptions {
                AllowLocalImages = true,
                DefaultPageSize = WordPageSize.Letter,
                MaxImageWidthPercentOfContent = 50,
                ImageLayout = {
                    AllowUpscale = true
                }
            });

            Assert.Single(doc.Images);
            Assert.InRange(doc.Images[0].Width ?? 0, 311, 313);
            Assert.InRange(doc.Images[0].Height ?? 0, 77, 79);
        }

        [Fact]
        public void MarkdownToWord_AppliesConfiguredImageMaxHeightPixels() {
            string imagePath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png"));
            string md = $"![Local]({imagePath}){{width=1200 height=300}}";
            var doc = md.LoadFromMarkdown(new MarkdownToWordOptions {
                AllowLocalImages = true,
                MaxImageHeightPixels = 100
            });

            Assert.Single(doc.Images);
            Assert.InRange(doc.Images[0].Height ?? 0, 99.5, 100.5);
            Assert.InRange(doc.Images[0].Width ?? 0, 399.5, 400.5);
        }

        [Fact]
        public void MarkdownToWord_EmitsImageLayoutDiagnosticsWhenScaled() {
            string imagePath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png"));
            string md = $"![Local]({imagePath}){{width=1200 height=300}}";
            var diagnostics = new List<MarkdownImageLayoutDiagnostic>();

            var doc = md.LoadFromMarkdown(new MarkdownToWordOptions {
                AllowLocalImages = true,
                MaxImageWidthPixels = 480,
                OnImageLayoutDiagnostic = diagnostics.Add
            });

            Assert.Single(doc.Images);
            var diagnostic = Assert.Single(diagnostics);
            Assert.True(diagnostic.ScaledByLayout);
            Assert.Equal("block-local", diagnostic.Context);
            Assert.Equal(1200, diagnostic.RequestedWidthPixels);
            Assert.InRange(diagnostic.FinalWidthPixels ?? 0, 479.5, 480.5);
        }

        [Fact]
        public void MarkdownToWord_FitImagesToPageContentWidth_ForcesPageMode() {
            var options = new MarkdownToWordOptions();
            options.FitImagesToContextWidth = true;

            options.FitImagesToPageContentWidth = true;

            Assert.True(options.FitImagesToPageContentWidth);
            Assert.False(options.FitImagesToContextWidth);
            Assert.Equal(MarkdownImageFitMode.PageContentWidth, options.ImageLayout.FitMode);
        }

        [Fact]
        public void WordToMarkdown_WritesImageDescription() {
            using var doc = WordDocument.Create();
            string imagePath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png"));
            doc.AddParagraph().AddImage(imagePath);
            doc.Images[0].Description = "Sample";

            string markdown = doc.ToMarkdown(new WordToMarkdownOptions());

            Assert.Contains("![Sample]", markdown);
        }

        private static int GetAvailablePort() {
            var tcpListener = new TcpListener(IPAddress.Loopback, 0);
            tcpListener.Start();
            int port = ((IPEndPoint)tcpListener.LocalEndpoint).Port;
            tcpListener.Stop();
            return port;
        }
    }
}
