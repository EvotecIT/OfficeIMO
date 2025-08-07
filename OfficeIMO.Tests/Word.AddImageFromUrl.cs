using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Threading.Tasks;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_AddImageFromUrl() {
            var filePath = Path.Combine(_directoryWithFiles, "ImageFromUrl.docx");
            string imagePath = Path.Combine(_directoryWithImages, "Kulek.jpg");

            int port = GetAvailablePort();

            using var listener = new HttpListener();
            listener.Prefixes.Add($"http://localhost:{port}/");
            listener.Start();
            var serverTask = Task.Run(() => {
                var context = listener.GetContext();
                var bytes = File.ReadAllBytes(imagePath);
                context.Response.ContentType = "image/jpeg";
                context.Response.ContentLength64 = bytes.Length;
                context.Response.OutputStream.Write(bytes, 0, bytes.Length);
                context.Response.OutputStream.Close();
                listener.Stop();
            });

            using (var document = WordDocument.Create(filePath)) {
                var img = document.AddImageFromUrl($"http://localhost:{port}/", 40, 40);
                Assert.NotNull(img);
                document.Save(false);
            }

            serverTask.Wait();

            using (var document = WordDocument.Load(filePath)) {
                Assert.Single(document.Images);
            }
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
