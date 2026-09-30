using OfficeIMO.Tool.Agent;
using OfficeIMO.Tool.Commands.Agent;
using System.Text.Json;
using Xunit;
using OfficeIMO.Email;
using OfficeIMO.Email.Store;
using System.IO.Compression;

namespace OfficeIMO.Tool.Tests;

public sealed class AgentEmailContentSearchTests {
    [Theory]
    [InlineData("directory-eml")]
    [InlineData("maildir-msg")]
    [InlineData("mbox")]
    [InlineData("emlx")]
    [InlineData("olm")]
    public async Task RequestedDecodeLimitAppliesBeforeContentSearchAcrossBackends(string kind) {
        string root = CreateMailbox(1);
        try {
            string input = root;
            if (kind == "maildir-msg") {
                File.Delete(Path.Combine(root, "000.eml"));
                var message = new EmailDocument { Subject = "decoded limit" };
                message.Body.Text = "body needle with more than one byte";
                Directory.CreateDirectory(Path.Combine(root, "new"));
                File.WriteAllBytes(Path.Combine(root, "new", "message"), message.ToBytes(EmailFileFormat.OutlookMsg));
            } else if (kind == "mbox") {
                input = Path.Combine(root, "message.mbox");
                File.WriteAllText(input, "From a@example.test Sat Jan 01 00:00:00 2022\nSubject: limit\n\nbody needle with more than one byte\n");
            } else if (kind == "emlx") {
                input = Path.Combine(root, "message.emlx");
                byte[] message = Encoding.UTF8.GetBytes("Subject: limit\r\n\r\nbody needle with more than one byte");
                File.WriteAllBytes(input, Encoding.ASCII.GetBytes(message.Length + "\n").Concat(message).ToArray());
            } else if (kind == "olm") {
                input = Path.Combine(root, "message.olm");
                using (var archive = new ZipArchive(File.Create(input), ZipArchiveMode.Create)) {
                    using var writer = new StreamWriter(archive.CreateEntry("Local/com.microsoft.__Messages/Inbox/Messages.xml").Open());
                    writer.Write("<emails><email><OPFMessageCopySubject>limit</OPFMessageCopySubject><OPFMessageCopyBody>body needle with more than one byte</OPFMessageCopyBody></email></emails>");
                }
            }
            var service = new OfficeImoAgentService(new AgentPathPolicy(new[] { root }));
            try {
                var page = await service.SearchEmailContentAsync(input, "body needle", fields: "TextBody", maxDecodedBytes: 1);
                Assert.Empty(page.Results);
                Assert.True(page.ItemsSkipped > 0 || page.DiagnosticCount > 0);
            } catch (EmailStoreLimitExceededException) { /* Eager store opening rejects before any response is produced. */ }
        } finally { Directory.Delete(root, recursive: true); }
    }
    [Fact]
    public async Task OutputTrimmedPagesResumeAfterDeliveredHitsAcrossServiceInstances() {
        string root = CreateMailbox(25);
        try {
            var ids = new HashSet<string>();
            string? checkpoint = null;
            bool shortened = false;
            do {
                var service = new OfficeImoAgentService(new AgentPathPolicy(new[] { root }));
                var page = await service.SearchEmailContentAsync(root, "body needle", fields: "TextBody",
                    checkpoint: checkpoint, take: 25, maxOutputCharacters: 700);
                Assert.True(AgentJson.Serialize(page).Length <= 700);
                Assert.NotEmpty(page.Results);
                foreach (var hit in page.Results) Assert.True(ids.Add(hit.Id), "Continuation repeated a delivered match.");
                shortened |= page.Returned < 25 && page.NextCheckpoint != null;
                checkpoint = page.NextCheckpoint;
                Assert.Equal(checkpoint == null, page.IsComplete);
                Assert.True(ids.Count <= 25);
            } while (checkpoint != null);
            Assert.True(shortened);
            Assert.Equal(25, ids.Count);
        } finally { Directory.Delete(root, recursive: true); }
    }

    [Fact]
    public async Task EmptyScanBatchHasDurableProgressAndRejectsChangedQueryAndSource() {
        string root = CreateMailbox(3);
        try {
            string first = Path.Combine(root, "000.eml");
            await File.WriteAllTextAsync(first, "Subject: First\r\n\r\nno match");
            var service = new OfficeImoAgentService(new AgentPathPolicy(new[] { root }));
            var page = await service.SearchEmailContentAsync(root, "body needle", fields: "TextBody", maxItemsScanned: 1);
            Assert.Empty(page.Results);
            Assert.Equal(1, page.ItemsScanned);
            Assert.True(page.ScanLimitReached);
            Assert.False(page.IsComplete);
            Assert.NotNull(page.NextCheckpoint);
            var resumed = await service.SearchEmailContentAsync(root, "body needle", fields: "TextBody",
                maxItemsScanned: 1, checkpoint: page.NextCheckpoint);
            Assert.Single(resumed.Results);
            Assert.Equal("TextBody", resumed.Results[0].MatchedFields);
            await Assert.ThrowsAsync<AgentUsageException>(() => service.SearchEmailContentAsync(root, "different",
                fields: "TextBody", maxItemsScanned: 1, checkpoint: page.NextCheckpoint));
            await File.AppendAllTextAsync(first, "changed");
            await Assert.ThrowsAsync<AgentUsageException>(() => service.SearchEmailContentAsync(root, "body needle",
                fields: "TextBody", maxItemsScanned: 1, checkpoint: page.NextCheckpoint));
        } finally { Directory.Delete(root, recursive: true); }
    }

    [Fact]
    public async Task CliSemanticSearchAndSelectiveFetchUseRealMailboxContent() {
        string root = CreateMailbox(2);
        try {
            using var output = new StringWriter(); using var error = new StringWriter();
            int exit = await AgentCommand.RunAsync(new[] { "search-email", root, "--query", "body needle", "--fields", "TextBody",
                "--take", "1", "--max-items-scanned", "2" }, output, error);
            Assert.Equal((int)OfficeImoToolExitCode.Success, exit);
            Assert.Equal(string.Empty, error.ToString());
            using var json = JsonDocument.Parse(output.ToString());
            Assert.Equal(1, json.RootElement.GetProperty("returned").GetInt32());
            Assert.True(json.RootElement.TryGetProperty("nextCheckpoint", out _));
            using var fetched = new StringWriter(); using var fetchError = new StringWriter();
            int fetchedExit = await AgentCommand.RunAsync(new[] { "fetch", "--path", root,
                "--source-id", json.RootElement.GetProperty("sourceId").GetString()!,
                "--id", json.RootElement.GetProperty("results")[0].GetProperty("id").GetString()! }, fetched, fetchError);
            Assert.Equal((int)OfficeImoToolExitCode.Success, fetchedExit);
            Assert.Contains("body needle 0", fetched.ToString());
            Assert.DoesNotContain("body needle 1", fetched.ToString());
        } finally { Directory.Delete(root, recursive: true); }
    }

    [Theory]
    [InlineData("search", "--fields", "All")]
    [InlineData("inspect", "--max-items-scanned", "100")]
    [InlineData("search-email", "--cursor", "1")]
    public void ContentOptionsCannotSilentlyReachAnotherCommand(string command, string option, string value) {
        Assert.Throws<AgentUsageException>(() => AgentArguments.Parse(new[] { command, "mailbox", "--query", "needle", option, value }));
    }

    private static string CreateMailbox(int count) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-content-agent-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        for (int index = 0; index < count; index++)
            File.WriteAllText(Path.Combine(root, index.ToString("D3") + ".eml"),
                "From: sender@example.test\r\nSubject: " + new string('s', 200) + "\r\n\r\nbody needle " + index);
        return root;
    }
}
