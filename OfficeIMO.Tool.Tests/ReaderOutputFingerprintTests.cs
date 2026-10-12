using System.Security.Cryptography;
using OfficeIMO.Tool.Commands.Reader;
using Xunit;

namespace OfficeIMO.Tool.Tests;

public sealed class ReaderOutputFingerprintTests {
    public static IEnumerable<object[]> Documents() {
        yield return [string.Empty];
        yield return ["# Café 😀\n\nUnpaired surrogate: \ud800"];
        yield return [new string('x', 4095) + "😀" + new string('é', 4096)];
    }

    [Theory]
    [MemberData(nameof(Documents))]
    public async Task Fingerprint_matches_persisted_utf8_bytes(string content) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-output-fingerprint-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            var fingerprint = ReaderToolOutput.FingerprintText(content, CancellationToken.None);
            string output = Path.Combine(root, "document.md");
            await ReaderToolOutput.WriteFileAsync(output, content, false, CancellationToken.None);
            byte[] bytes = File.ReadAllBytes(output);
            Assert.Equal(bytes.LongLength, fingerprint.LengthBytes);
            Assert.Equal(Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant(), fingerprint.Sha256);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }
}
