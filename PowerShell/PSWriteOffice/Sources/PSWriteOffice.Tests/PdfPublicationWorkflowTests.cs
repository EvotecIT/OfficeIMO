using System.Collections.ObjectModel;
using System.Globalization;
using System.Management.Automation;
using System.Management.Automation.Host;
using System.Management.Automation.Runspaces;
using System.Security;
using OfficeIMO;
using OfficeIMO.Pdf;
using PSWriteOffice.Cmdlets.Conversion;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class PdfPublicationWorkflowTests {
    [Theory]
    [InlineData("Optimize-OfficePdf")]
    [InlineData("Remove-OfficePdfContent")]
    public void NoClobberRetainsADestinationCreatedDuringNativeConfirmation(string command) {
        string root = Path.Combine(AppContext.BaseDirectory, "pdf-confirmation-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string source = Path.Combine(root, "source.pdf"), destination = Path.Combine(root, "result.pdf");
            using (var pdf = PdfDocument.Create(document => document.Content(content => content.AddParagraph("account:123")))) pdf.Save(source);
            byte[] sentinel = [4, 8, 15, 16, 23, 42];
            DateTime timestamp = new(2025, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            int prompts = 0;
            var host = new ConfirmationHost(() => {
                Assert.False(File.Exists(destination));
                prompts++;
                File.WriteAllBytes(destination, sentinel);
                File.SetLastWriteTimeUtc(destination, timestamp);
            });
            var state = InitialSessionState.CreateDefault();
            state.ImportPSModule(new[] { typeof(ConvertOfficeDocumentCommand).Assembly.Location });
            using var runspace = RunspaceFactory.CreateRunspace(host, state);
            runspace.Open();
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddCommand(command).AddParameter("LiteralPath", source).AddParameter("OutputPath", destination)
                .AddParameter("NoClobber").AddParameter("Confirm").AddParameter("ErrorAction", "Stop");
            if (command == "Remove-OfficePdfContent") shell.AddParameter("Pattern", "account:\\d+");
            Assert.ThrowsAny<RuntimeException>(() => shell.Invoke());
            Assert.Equal(1, prompts);
            Assert.Equal(sentinel, File.ReadAllBytes(destination));
            Assert.Equal(timestamp, File.GetLastWriteTimeUtc(destination));
            using var preserved = PdfDocument.Load(source);
            Assert.Contains("account:123", preserved.Read().Text);
        } finally { Directory.Delete(root, true); }
    }

    private sealed class ConfirmationHost(Action confirm) : PSHost {
        private readonly PSHostUserInterface _ui = new ConfirmationUi(confirm);
        public override Guid InstanceId { get; } = Guid.NewGuid();
        public override string Name => "Native confirmation contract";
        public override Version Version => new(1, 0);
        public override PSHostUserInterface UI => _ui;
        public override CultureInfo CurrentCulture => CultureInfo.InvariantCulture;
        public override CultureInfo CurrentUICulture => CultureInfo.InvariantCulture;
        public override void EnterNestedPrompt() => throw new NotSupportedException();
        public override void ExitNestedPrompt() => throw new NotSupportedException();
        public override void NotifyBeginApplication() { }
        public override void NotifyEndApplication() { }
        public override void SetShouldExit(int exitCode) { }
    }

    private sealed class ConfirmationUi(Action confirm) : PSHostUserInterface {
        public override PSHostRawUserInterface RawUI => null!;
        public override int PromptForChoice(string caption, string message, Collection<ChoiceDescription> choices, int defaultChoice) {
            confirm();
            return 0;
        }
        public override Dictionary<string, PSObject> Prompt(string caption, string message, Collection<FieldDescription> descriptions) => throw new NotSupportedException();
        public override PSCredential PromptForCredential(string caption, string message, string userName, string targetName) => throw new NotSupportedException();
        public override PSCredential PromptForCredential(string caption, string message, string userName, string targetName,
            PSCredentialTypes allowedCredentialTypes, PSCredentialUIOptions options) => throw new NotSupportedException();
        public override string ReadLine() => throw new NotSupportedException();
        public override SecureString ReadLineAsSecureString() => throw new NotSupportedException();
        public override void Write(string value) { }
        public override void Write(ConsoleColor foregroundColor, ConsoleColor backgroundColor, string value) { }
        public override void WriteLine(string value) { }
        public override void WriteErrorLine(string value) { }
        public override void WriteDebugLine(string message) { }
        public override void WriteProgress(long sourceId, ProgressRecord record) { }
        public override void WriteVerboseLine(string message) { }
        public override void WriteWarningLine(string message) { }
    }
}
