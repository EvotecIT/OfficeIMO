using System;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Threading.Tasks;
using SixLabors.ImageSharp;
using OfficeIMO.Word.Markdown;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Markdown {
        [Fact]
        public void MarkdownToWord_ParsesImageHints() {
            string imagePath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png"));
            int port = GetAvailablePort();

            using var listener = new HttpListener();
            listener.Prefixes.Add($"http://localhost:{port}/");
            listener.Start();
            var serverTask = Task.Run(() => {
                var context = listener.GetContext();
                var bytes = File.ReadAllBytes(imagePath);
                context.Response.ContentType = "image/png";
                context.Response.ContentLength64 = bytes.Length;
                context.Response.OutputStream.Write(bytes, 0, bytes.Length);
                context.Response.OutputStream.Flush();
                context.Response.Close();
            });

            string md = $"![Local]({imagePath} \"Desc local\"){{width=40 height=30}}\n" +
                         $"![Remote](http://localhost:{port}/ \"Desc remote\"){{width=50 height=20}}";
            var doc = md.LoadFromMarkdown(new MarkdownToWordOptions());

            Assert.Equal(2, doc.Images.Count);
            Assert.Equal("Desc local", doc.Images[0].Description);
            Assert.Equal(40, doc.Images[0].Width);
            Assert.Equal(30, doc.Images[0].Height);
            Assert.Equal("Desc remote", doc.Images[1].Description);
            Assert.Equal(50, doc.Images[1].Width);
            Assert.Equal(20, doc.Images[1].Height);

            serverTask.Wait();
            listener.Stop();
        }

        [Fact]
        public void MarkdownToWord_UsesNaturalSizeWhenNoHints() {
            string imagePath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png"));
            string md = $"![Local]({imagePath})";
            var doc = md.LoadFromMarkdown(new MarkdownToWordOptions());

            using var image = Image.Load(imagePath, out _);

            Assert.Single(doc.Images);
            Assert.Equal(image.Width, doc.Images[0].Width);
            Assert.Equal(image.Height, doc.Images[0].Height);
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
