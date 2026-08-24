using OfficeIMO.GoogleWorkspace;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class GoogleWorkspaceDiagnosticsTests {
        [Fact]
        public async Task Test_GoogleWorkspaceHttpTransport_DoesNotPreallocateUntrustedDeclaredLength() {
            byte[] expected = Encoding.UTF8.GetBytes("small response");
            using var httpClient = new HttpClient(new FakeHttpMessageHandler(_ =>
                Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) {
                    Content = new MisreportedLengthContent(expected, int.MaxValue)
                })));
            using var transport = new GoogleWorkspaceHttpTransport(
                new GoogleWorkspaceSessionOptions {
                    HttpClient = httpClient,
                    MaxRetryCount = 0
                });

            byte[] actual = await transport.SendBytesAsync(
                "token",
                HttpMethod.Get,
                "https://www.googleapis.com/drive/v3/files/file-1?alt=media",
                GoogleWorkspaceRequestSafety.Safe,
                "Google Drive API",
                new TranslationReport());

            Assert.Equal(expected, actual);
        }

        [Fact]
        public async Task Test_GoogleWorkspaceHttpTransport_TruncatesUnderreportedErrorResponses() {
            byte[] responseBytes = Encoding.UTF8.GetBytes(new string('x', 128 * 1024));
            using var httpClient = new HttpClient(new FakeHttpMessageHandler(_ =>
                Task.FromResult(new HttpResponseMessage(HttpStatusCode.BadRequest) {
                    Content = new MisreportedLengthContent(responseBytes, declaredLength: 1)
                })));
            using var transport = new GoogleWorkspaceHttpTransport(
                new GoogleWorkspaceSessionOptions {
                    HttpClient = httpClient,
                    MaxRetryCount = 0
                });

            GoogleWorkspaceApiException exception =
                await Assert.ThrowsAsync<GoogleWorkspaceApiException>(() =>
                    transport.SendBytesAsync(
                        "token",
                        HttpMethod.Get,
                        "https://lh3.googleusercontent.com/image.png",
                        GoogleWorkspaceRequestSafety.Safe,
                        "Google content",
                        new TranslationReport()));

            Assert.Equal(64 * 1024, exception.ResponseBody.Length);
        }

        private sealed class MisreportedLengthContent : HttpContent {
            private readonly byte[] _bytes;
            private readonly long _declaredLength;

            internal MisreportedLengthContent(byte[] bytes, long declaredLength) {
                _bytes = bytes;
                _declaredLength = declaredLength;
            }

            protected override Task SerializeToStreamAsync(Stream stream, TransportContext? context) =>
                throw new InvalidOperationException(
                    "ResponseContentRead attempted to buffer the response.");

            protected override Task<Stream> CreateContentReadStreamAsync() =>
                Task.FromResult<Stream>(new MemoryStream(_bytes, writable: false));

            protected override bool TryComputeLength(out long length) {
                length = _declaredLength;
                return true;
            }
        }
    }
}
