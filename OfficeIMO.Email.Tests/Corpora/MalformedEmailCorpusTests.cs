using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.Email.AddressBook;
using OfficeIMO.Email.Store;

namespace OfficeIMO.Email.Tests;

public sealed class MalformedEmailCorpusTests {
    [Fact]
    public void ManifestHashesAndMalformedCasesRemainBoundedAndCancellationAware() {
        string root = FindRepositoryRoot();
        using JsonDocument manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(root,
            "OfficeIMO.Email.Tests", "Corpora", "malformed-email-corpus.json")));
        foreach (JsonElement item in manifest.RootElement.GetProperty("cases").EnumerateArray()) {
            byte[] bytes = Convert.FromBase64String(item.GetProperty("base64").GetString()!);
            Assert.Equal(item.GetProperty("sha256").GetString(), Hash(bytes));
            string format = item.GetProperty("format").GetString()!;
            string expected = item.GetProperty("expected").GetString()!;
            string actual = Exercise(format, bytes, CancellationToken.None);
            Assert.True(string.Equals(expected, actual, StringComparison.Ordinal),
                $"Malformed corpus result changed for {item.GetProperty("id").GetString()}: expected {expected}, actual {actual}.");
            using var canceled = new CancellationTokenSource(); canceled.Cancel();
            Exception? cancellation = Record.Exception(() => Exercise(format, bytes, canceled.Token));
            Assert.True(cancellation is OperationCanceledException,
                $"Malformed corpus cancellation was not observed by {item.GetProperty("id").GetString()}; actual exception: {cancellation?.GetType().Name ?? "none"}.");
        }
        using JsonDocument producers = JsonDocument.Parse(File.ReadAllText(Path.Combine(root,
            "OfficeIMO.Email.Tests", "Corpora", "producer-corpora.json")));
        Assert.All(producers.RootElement.GetProperty("sources").EnumerateArray(), source => {
            Assert.False(string.IsNullOrWhiteSpace(source.GetProperty("producer").GetString()));
            Assert.Equal("live-opt-in", source.GetProperty("executionStatus").GetString());
            Assert.NotEmpty(source.GetProperty("requiredFormats").EnumerateArray());
            Assert.Contains(source.GetProperty("artifactEvidence").EnumerateArray(), field =>
                string.Equals(field.GetString(), "sha256", StringComparison.Ordinal));
            Assert.False(string.IsNullOrWhiteSpace(source.GetProperty("oracle").GetString()));
        });
    }

    private static string Exercise(string format, byte[] bytes, CancellationToken token) {
        try {
            switch (format) {
                case "eml": case "msg": case "tnef":
                    using (EmailReadResult result = new EmailDocumentReader().Read(bytes, token))
                        return result.Diagnostics.Any(diagnostic => diagnostic.Severity == EmailDiagnosticSeverity.Error)
                            ? "diagnostic-or-rejection" : "bounded-result";
                case "pst": case "ost":
                    using (var stream = new MemoryStream(bytes, false))
                        return new EmailStoreReader().Read(stream, format + ".pst", token).HasErrors
                            ? "diagnostic-or-rejection" : "bounded-result";
                case "olm":
                    using (var stream = new MemoryStream(bytes, false))
                        return new EmailStoreReader().Read(stream, "broken.olm", token).HasErrors
                            ? "diagnostic-or-rejection" : "bounded-result";
                case "oab":
                    using (var stream = new MemoryStream(bytes, false))
                    using (OfflineAddressBookSession.Open(stream, "broken.oab", cancellationToken: token))
                        return "bounded-result";
                case "ics":
                    using (var stream = new MemoryStream(bytes, false))
                        return IcsDocument.Load(stream, cancellationToken: token).Validate().Count > 0
                            ? "validation-result" : "bounded-result";
                case "vcf":
                    using (var stream = new MemoryStream(bytes, false))
                        return VCardDocument.Load(stream, cancellationToken: token).Validate().Count > 0
                            ? "validation-result" : "bounded-result";
                default: throw new InvalidDataException("Unknown corpus format: " + format);
            }
        } catch (Exception exception) when (exception is InvalidDataException || exception is EndOfStreamException ||
            exception is NotSupportedException || exception is EmailStoreLimitExceededException ||
            exception is OfflineAddressBookLimitExceededException) {
            return "diagnostic-or-rejection";
        }
    }

    private static string Hash(byte[] bytes) { using SHA256 hash = SHA256.Create(); return BitConverter.ToString(hash.ComputeHash(bytes)).Replace("-", string.Empty).ToLowerInvariant(); }
    private static string FindRepositoryRoot() { DirectoryInfo? directory = new DirectoryInfo(AppContext.BaseDirectory); while (directory != null && !File.Exists(Path.Combine(directory.FullName, "OfficeIMO.sln"))) directory = directory.Parent; return directory?.FullName ?? throw new DirectoryNotFoundException(); }
}
