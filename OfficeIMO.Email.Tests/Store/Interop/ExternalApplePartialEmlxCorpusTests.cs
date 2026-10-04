using OfficeIMO.Email;
using System.Security.Cryptography;

namespace OfficeIMO.Email.Store.Tests;

public sealed class ExternalApplePartialEmlxCorpusTests {
    // Pinned independent fixture: qqilihq/partial-emlx-converter, MIT, commit 1263a7326c3078ba2e86c76f26dcc2d661f0adcf.
    [EnvironmentFact("OFFICEIMO_EMAIL_APPLE_PARTIAL_CORPUS", requireDirectory: true)]
    public void RecoversPinnedIndependentSiblingAndPreservesDecodedBytesThroughEmlExport() {
        string root = Environment.GetEnvironmentVariable("OFFICEIMO_EMAIL_APPLE_PARTIAL_CORPUS")!;
        var files = new Dictionary<string, string> {
            ["Messages/114892.partial.emlx"] = "e6ac7d446a7b0be03fefd2c077a3387277ccac709049be9251a337d949050d15",
            ["Attachments/114892/2.2/short.txt"] = "50ffb4ec5d05f84df226ecde9869ebdcdd8937d736688d49948cf636a3f22ca4",
            ["Attachments/114892/2.4/original.doc"] = "e28d914a7bb97863292f361646ee60ca7ba340ab1d3b96e5201b8a23b27312af",
            ["Attachments/114892/2.6/text.txt"] = "7061027a4c13369d5543bbe7b9cf4f7125043a9a06b17ac3f3f680f9968cd771",
            ["Attachments/114892/2.8/image001.png"] = "a3c35e34cbdd1100e35c1a8dfe1d6937974483af8f2e710458894b818dafa309"
        };
        foreach (KeyValuePair<string, string> file in files) Assert.Equal(file.Value, Hash(File.ReadAllBytes(Path.Combine(root, file.Key))));
        string[] expected = files.Where(file => file.Key.StartsWith("Attachments/", StringComparison.Ordinal)).Select(file => file.Value).ToArray();
        using (EmailStoreSession session = EmailStoreSession.Open(root)) {
            EmailStoreItem item = session.ReadItem(Assert.Single(session.EnumerateItems()),
                new EmailStoreItemReadOptions(preferStreamingAttachmentContent: true));
            Assert.Equal(4, item.Document.Properties["Emlx:RecoveredPartCount"]);
            Assert.Equal(0, item.Document.Properties["Emlx:UnresolvedPartCount"]);
            var payloads = new HashSet<string>(StringComparer.Ordinal);
            foreach (EmailAttachment recovered in item.Document.Attachments) {
                using Stream input = recovered.OpenContentStream();
                using var output = new MemoryStream();
                input.CopyTo(output);
                string hash = Hash(output.ToArray());
                payloads.Add(hash);
                if (expected.Contains(hash)) {
                    Assert.Null(recovered.Content);
                    Assert.NotNull(recovered.ContentSource);
                    Assert.Equal(output.Length, recovered.Length);
                }
            }
            Assert.All(expected, hash => Assert.Contains(hash, payloads));
            using EmailReadResult exported = new EmailDocumentReader().Read(
                new EmailDocumentWriter(new EmailWriterOptions(conversionLossPolicy: EmailConversionLossPolicy.Warn)).ToBytes(item.Document, EmailFileFormat.Eml));
            string[] exportedHashes = exported.Document.Attachments.Select(attachment => Hash(attachment.Content!)).ToArray();
            Assert.All(expected, hash => Assert.Contains(hash, exportedHashes));
        }
        foreach (KeyValuePair<string, string> file in files) Assert.Equal(file.Value, Hash(File.ReadAllBytes(Path.Combine(root, file.Key))));
    }

    private static string Hash(byte[] bytes) {
        using SHA256 hash = SHA256.Create();
        return string.Concat(hash.ComputeHash(bytes).Select(value => value.ToString("x2", System.Globalization.CultureInfo.InvariantCulture)));
    }
}
