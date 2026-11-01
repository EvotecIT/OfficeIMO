using System.Management.Automation;
using System.Management.Automation.Runspaces;
using OfficeIMO.Pdf;
using OfficeIMO.Ocr.Process;
using PSWriteOffice.Cmdlets.Reader;
using PSWriteOffice.Models.Pdf;
using PSWriteOffice.Models.Reader;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class ReaderPdfPolicyWorkflowTests {
    [Fact]
    public void ReadFactorySnapshotsBasePolicyAndSelectsRequestedPages() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-pdf-read-factory-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string path = Path.Combine(root, "source.pdf");
            using (var document = PdfDocument.Create(pdf => {
                pdf.Page(page => page.Content(content => content.Text("First omitted page")));
                pdf.Page(page => page.Content(content => content.Text("Second selected page")));
            })) document.Save(path);
            var policy = new PdfLoadOptions {
                ParsingMode = PdfParsingMode.Strict, Password = "captured-test-policy",
                PermissionPolicy = PdfPermissionPolicy.IgnoreRestrictions,
                Limits = new PdfReadLimits { MaxInputBytes = new FileInfo(path).Length + 1 }
            };
            var layout = new PdfTextLayoutOptions { ForceSingleColumn = true };
            using var runspace = OpenRunspace();
            runspace.SessionStateProxy.SetVariable("policy", policy);
            runspace.SessionStateProxy.SetVariable("layout", layout);
            using var shell = PowerShell.Create(); shell.Runspace = runspace;
            shell.AddScript("New-OfficePdfReadOptions -LoadOptions $policy -Profile Fast -Page 2 -LayoutOptions $layout -IgnorePermissionRestrictions:$false");
            var options = Assert.IsType<PdfReadOptions>(Assert.Single(shell.Invoke()).BaseObject);
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            Assert.NotSame(policy, options.LoadOptions);
            Assert.NotSame(policy.Limits, options.LoadOptions.Limits);
            Assert.NotSame(layout, options.LoadOptions.LayoutOptions);
            Assert.Equal(PdfReadProfile.Fast, options.LoadOptions.Profile);
            Assert.Equal(PdfParsingMode.Strict, options.LoadOptions.ParsingMode);
            Assert.Equal(policy.Limits.MaxInputBytes, options.LoadOptions.Limits.MaxInputBytes);
            Assert.Equal(PdfPermissionPolicy.Enforce, options.LoadOptions.PermissionPolicy);
            Assert.Equal(PdfReadProfile.Structured, policy.Profile);
            Assert.Equal(PdfPermissionPolicy.IgnoreRestrictions, policy.PermissionPolicy);
            Assert.Null(policy.PageSelection);
            layout.ForceSingleColumn = false;
            Assert.True(options.LoadOptions.LayoutOptions.ForceSingleColumn);
            shell.Commands.Clear(); shell.Streams.Error.Clear();
            shell.AddCommand("Read-OfficeDocument").AddParameter("LiteralPath", path).AddParameter("Options", new object[] { options })
                .AddParameter("As", "Text").AddParameter("ErrorAction", "Stop");
            string text = string.Join("\n", shell.Invoke().Select(item => item.ToString()));
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            Assert.Contains("Second selected page", text);
            Assert.DoesNotContain("First omitted page", text);
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void OcrReadFactoryAppliesNativeRasterBudgetBeforeStartingAProvider() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-pdf-ocr-policy-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string path = Path.Combine(root, "source.pdf");
            using (var document = PdfDocument.Create(pdf => pdf.Page(page => page.Content(content => content.Text("OCR policy source")))))
                document.Save(path);
            var provider = new OfficeOcrOptions {
                Provider = new ProcessOcrEngine(new ProcessOcrEngineOptions { FileName = Path.Combine(root, "not-a-provider") })
            };
            using var runspace = OpenRunspace();
            runspace.SessionStateProxy.SetVariable("provider", provider);
            runspace.SessionStateProxy.SetVariable("path", path);
            using var shell = PowerShell.Create(); shell.Runspace = runspace;
            shell.AddScript("$settings = New-OfficePdfReadOptions -OcrMergeOptions @{ MaxPixelsPerPage = 1; Dpi = 72 }; try { Read-OfficeDocument -LiteralPath $path -As Ocr -Options $settings -OcrOptions $provider -ErrorAction Stop } catch { $_.Exception.Message }");
            string message = Assert.IsType<string>(Assert.Single(shell.Invoke()).BaseObject);
            Assert.Contains("pixel", message, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("not-a-provider", message, StringComparison.OrdinalIgnoreCase);
            shell.Commands.Clear(); shell.Streams.Error.Clear();
            shell.AddScript("try { Read-OfficeDocument -LiteralPath $path -As Text -Options $settings -OcrOptions $provider -ErrorAction Stop } catch { $_.Exception.Message }");
            Assert.Contains("As Ocr", Assert.IsType<string>(Assert.Single(shell.Invoke()).BaseObject), StringComparison.Ordinal);
            Assert.Equal(new[] { path }, Directory.GetFiles(root));
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void TypedPdfPolicyEnforcesInputAdmissionBeforeReadingText() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-reader-pdf-policy-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string path = Path.Combine(root, "source.pdf");
            using (var document = PdfDocument.Create()) {
                document.Compose(builder => builder.Content(content => content.Text("Reader policy source text")));
                document.Save(path);
            }
            var state = InitialSessionState.CreateDefault();
            state.ImportPSModule(new[] { typeof(ReadOfficeDocumentCommand).Assembly.Location });
            using var runspace = RunspaceFactory.CreateRunspace(state); runspace.Open();
            using var shell = PowerShell.Create(); shell.Runspace = runspace;
            shell.AddCommand("Read-OfficeDocument").AddParameter("LiteralPath", path).AddParameter("As", "Text")
                .AddParameter("Options", new object[] { new PdfLoadOptions {
                    Limits = new PdfReadLimits { MaxInputBytes = new FileInfo(path).Length + 1 }
                } }).AddParameter("ErrorAction", "Stop");
            var accepted = shell.Invoke();
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            Assert.Contains("Reader policy source text", string.Join("\n", accepted.Select(item => item.ToString())));
            shell.Commands.Clear(); shell.Streams.Error.Clear();
            shell.AddCommand("Read-OfficeDocument").AddParameter("LiteralPath", path).AddParameter("As", "Text")
                .AddParameter("Options", new object[] { new PdfLoadOptions {
                    Limits = new PdfReadLimits { MaxInputBytes = 1 }
                } }).AddParameter("ErrorAction", "Stop");
            Exception? failure = Record.Exception(() => shell.Invoke());
            var errors = shell.Streams.Error.Select(error => error.Exception).ToList();
            if (failure != null) errors.Add(failure);
            Assert.Contains(errors.SelectMany(ExceptionChain), error => error is PdfReadLimitException limit &&
                limit.Kind == PdfReadLimitKind.InputBytes && limit.Limit == 1);
        } finally { Directory.Delete(root, true); }
    }

    private static IEnumerable<Exception> ExceptionChain(Exception error) {
        yield return error;
        if (error is AggregateException aggregate) {
            foreach (var inner in aggregate.InnerExceptions)
                foreach (var nested in ExceptionChain(inner)) yield return nested;
        } else if (error.InnerException != null) {
            foreach (var nested in ExceptionChain(error.InnerException)) yield return nested;
        }
    }

    private static Runspace OpenRunspace() {
        var state = InitialSessionState.CreateDefault();
        state.ImportPSModule(new[] { typeof(ReadOfficeDocumentCommand).Assembly.Location });
        var runspace = RunspaceFactory.CreateRunspace(state); runspace.Open(); return runspace;
    }
}
