using System.Management.Automation;
using System.Management.Automation.Runspaces;
using OfficeIMO;
using OfficeIMO.Excel;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using PSWriteOffice.Cmdlets.Conversion;
using PSWriteOffice.Cmdlets.Excel;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class DocumentValidationWorkflowTests {
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void EncryptedPdfValidationCarriesOriginalAuthenticationWithoutChangingEitherReadPolicy(bool usePdfRead) {
        string root = Directory.CreateTempSubdirectory("office-validation-encrypted-").FullName;
        try {
            string sourcePath = Path.Combine(root, "source.pdf");
            string differencePath = Path.Combine(root, "difference.pdf");
            foreach (var target in new[] { (sourcePath, "source-reader", "source-owner"), (differencePath, "difference-reader", "difference-owner") }) {
                using var generated = PdfDocument.Create();
                generated.Compose(builder => builder.Content(content => content.AddParagraph("Protected validation content")));
                generated.Save(target.Item1, new PdfSaveOptions { AllowLoss = true, Encryption = new PdfStandardEncryptionOptions(target.Item2) {
                    OwnerPassword = target.Item3, Algorithm = PdfStandardEncryptionAlgorithm.Aes128, AllowedPermissions = PdfStandardPermissions.Print
                } });
            }
            byte[] sourceBytes = File.ReadAllBytes(sourcePath), differenceBytes = File.ReadAllBytes(differencePath);
            var originalRead = new PdfLoadOptions {
                PermissionPolicy = PdfPermissionPolicy.IgnoreRestrictions,
                Limits = new PdfReadLimits { MaxInputBytes = 2 * 1024 * 1024, MaxIndirectObjects = 5_000 }
            };
            var rewrittenRead = new PdfLoadOptions {
                Password = "difference-owner", PermissionPolicy = PdfPermissionPolicy.Enforce,
                Limits = new PdfReadLimits { MaxInputBytes = 3 * 1024 * 1024, MaxIndirectObjects = 6_000 }
            };
            var rewrite = new PdfRewritePreservationOptions {
                OriginalReadOptions = originalRead, RewrittenReadOptions = rewrittenRead, PreserveSecurityState = false
            }.RequireTextMarkers("Protected validation content");
            var pdfChecks = new OfficePdfValidationOptions { Rewrite = rewrite };
            var nativeOptions = new OfficeDocumentValidationOptions {
                Password = "source-reader", PdfRead = usePdfRead ? originalRead : null, Pdf = pdfChecks
            };
            var structure = OfficeDocument.Validate(sourcePath, options: nativeOptions);
            Assert.True(structure.IsValid, structure.ToMarkdown());
            Assert.True(Assert.Single(structure.NativeReports.OfType<PdfDocumentPreflight>()).CanReadLogicalObjects);
            pdfChecks.DifferencePath = differencePath;
            Assert.True(Assert.Single(OfficeDocument.Validate(sourcePath, options: nativeOptions).NativeReports.OfType<PdfRewritePreservationReport>()).IsPreserved);

            using var difference = PdfDocument.Load(differencePath, rewrittenRead);
            pdfChecks.DifferencePath = null;
            pdfChecks.DifferenceDocument = difference;
            Assert.True(Assert.Single(OfficeDocument.Validate(sourcePath, options: nativeOptions).NativeReports.OfType<PdfRewritePreservationReport>()).IsPreserved);
            var authenticatedRead = originalRead.Clone(null, "source-reader");
            using var borrowed = PdfDocument.Load(sourcePath, authenticatedRead);
            nativeOptions.Password = null;
            Assert.True(Assert.Single(OfficeDocument.Validate(borrowed, options: nativeOptions).NativeReports.OfType<PdfRewritePreservationReport>()).IsPreserved);

            pdfChecks.DifferenceDocument = null;
            pdfChecks.DifferencePath = differencePath;
            using var runspace = CreateRunspace();
            runspace.SessionStateProxy.SetVariable("source", sourcePath);
            runspace.SessionStateProxy.SetVariable("checks", pdfChecks);
            runspace.SessionStateProxy.SetVariable("read", originalRead);
            runspace.SessionStateProxy.SetVariable("authenticatedRead", authenticatedRead);
            runspace.SessionStateProxy.SetVariable("usePdfRead", usePdfRead);
            using var password = new System.Security.SecureString();
            foreach (char character in "source-reader") password.AppendChar(character);
            password.MakeReadOnly();
            runspace.SessionStateProxy.SetVariable("credential", new PSCredential("synthetic", password));
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddScript("""
                $options = @($checks)
                if ($usePdfRead) { $options += $read }
                Test-OfficeDocument -LiteralPath $source -Password $credential -Options $options
                Test-OfficeDocument -LiteralPath $source -Password $credential -Options $options -Quiet
                Test-OfficeDocument -LiteralPath $source -Options $checks,$authenticatedRead -Quiet
                """);
            var output = shell.Invoke();
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            Assert.Equal(3, output.Count);
            var report = Assert.IsType<OfficeValidationReport>(output[0].BaseObject);
            Assert.True(report.IsValid, report.ToMarkdown());
            Assert.Equal(PdfPermissionPolicy.IgnoreRestrictions, Assert.Single(report.NativeReports.OfType<PdfDocumentPreflight>()).PermissionPolicy);
            Assert.True(Assert.Single(report.NativeReports.OfType<PdfRewritePreservationReport>()).CanReadDetails);
            Assert.True(Assert.IsType<bool>(output[1].BaseObject));
            Assert.True(Assert.IsType<bool>(output[2].BaseObject));
            Assert.Same(originalRead, rewrite.OriginalReadOptions);
            Assert.Same(rewrittenRead, rewrite.RewrittenReadOptions);
            Assert.Null(originalRead.Password);
            Assert.Equal("difference-owner", rewrittenRead.Password);
            Assert.Equal("source-reader", authenticatedRead.Password);
            Assert.Equal(5_000, originalRead.Limits.MaxIndirectObjects);
            Assert.Equal(6_000, rewrittenRead.Limits.MaxIndirectObjects);
            Assert.Equal(PdfPermissionPolicy.IgnoreRestrictions, originalRead.PermissionPolicy);
            Assert.Equal(PdfPermissionPolicy.Enforce, rewrittenRead.PermissionPolicy);
            Assert.Single(rewrite.RequiredTextMarkers);
            Assert.False(rewrite.PreserveSecurityState);
            Assert.Equal(sourceBytes, File.ReadAllBytes(sourcePath));
            Assert.Equal(differenceBytes, File.ReadAllBytes(differencePath));
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void PdfComplianceHashtableProfileReachesTheComplianceValidator() {
        using var pdf = PdfDocument.Create();
        pdf.Compose(builder => builder.Content(content => content.AddParagraph("Compliance")));
        using var runspace = CreateRunspace();
        runspace.SessionStateProxy.SetVariable("pdf", pdf);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript("Test-OfficeDocument -Document $pdf -Check Compliance -Options @{ Profile = 'PdfA2B' }");
        var output = shell.Invoke();
        Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
        var report = Assert.IsType<OfficeValidationReport>(Assert.Single(output).BaseObject);
        Assert.Equal(PdfComplianceProfile.PdfA2B, Assert.Single(report.NativeReports.OfType<PdfComplianceReadinessReport>()).Profile);
    }

    [Fact]
    public void FileInfoLiteralAndWildcardInputsReturnReportsOrQuietBooleansWithoutChangingTheFile() {
        string root = Directory.CreateTempSubdirectory("office-validation-binding-").FullName;
        try {
            string path = Path.Combine(root, "proposal[draft].docx");
            using (var word = WordDocument.Create()) { word.AddParagraph("Proposal"); word.Save(path); }
            byte[] before = File.ReadAllBytes(path);
            using var runspace = CreateRunspace();
            runspace.SessionStateProxy.SetVariable("path", path);
            runspace.SessionStateProxy.SetVariable("root", root);
            runspace.SessionStateProxy.SetVariable("policy", OfficePackageSecurityOptions.SecureDefaults);
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddScript("""
                [IO.FileInfo]::new($path) | Test-OfficeDocument -Check Security -Options @{ MaxPackageBytes = 1 } -Quiet
                Test-OfficeDocument -LiteralPath $path -Check Security -Options $policy
                Test-OfficeDocument -Path (Join-Path $root '*.docx') -Check Schema
                """);
            var output = shell.Invoke();
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            Assert.Equal(3, output.Count);
            Assert.False(Assert.IsType<bool>(output[0].BaseObject));
            Assert.True(Assert.IsType<OfficeValidationReport>(output[1].BaseObject).IsValid);
            Assert.IsType<OfficeValidationReport>(output[2].BaseObject);
            Assert.Equal(before, File.ReadAllBytes(path));
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void LiveBindingConsumesExcelTestOptionsAndTheReadOnlyCommandRejectsWhatIfAndDocumentPasswords() {
        using var workbook = ExcelDocument.Create();
        workbook.AddSheet("Data").Cell(1, 1).SetValue(3);
        using var runspace = CreateRunspace();
        runspace.SessionStateProxy.SetVariable("book", workbook);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript("""
            $checks = New-OfficeExcelTestOptions -Capability ReadWorkbookData -IncludeFeatures -IncludeRepairHints -NoOpenXmlValidation
            $book | Test-OfficeDocument -Options $checks
            Test-OfficeDocument -Document $book -Check Structure -Options @{ ValidateOpenXml = $false } -Quiet
            """);
        var output = shell.Invoke();
        Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
        Assert.Equal(2, output.Count);
        var report = Assert.IsType<OfficeValidationReport>(output[0].BaseObject);
        Assert.Contains(report.NativeReports, evidence => evidence is ExcelFeatureReport);
        Assert.IsType<bool>(output[1].BaseObject);
        foreach (string arguments in new[] { "-WhatIf", "-Password 'secret'" }) {
            shell.Commands.Clear();
            shell.Streams.Error.Clear();
            shell.AddScript("Test-OfficeDocument -Document $book " + arguments);
            Assert.Empty(shell.Invoke());
            Assert.True(shell.HadErrors);
            Assert.Contains(shell.Streams.Error, error => error.CategoryInfo.Category == ErrorCategory.InvalidArgument);
        }
        workbook.Sheets[0].Cell(2, 1).SetValue("Still editable");
    }

    [Fact]
    public void PathPasswordsUseTheSharedTransformationAndWarnOnlyForPlainText() {
        string root = Directory.CreateTempSubdirectory("office-validation-password-").FullName;
        try {
            string path = Path.Combine(root, "document.docx");
            using (var word = WordDocument.Create()) { word.AddParagraph("Password binding"); word.Save(path); }
            using var runspace = CreateRunspace();
            runspace.SessionStateProxy.SetVariable("path", path);
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddScript("Test-OfficeDocument -LiteralPath $path -Check Security -Password 'secret' -Quiet");
            Assert.True(Assert.IsType<bool>(Assert.Single(shell.Invoke()).BaseObject));
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            Assert.Single(shell.Streams.Warning);
            shell.Commands.Clear();
            shell.Streams.Warning.Clear();
            shell.AddScript("""
                $credential = [pscredential]::new('user', (ConvertTo-SecureString 'secret' -AsPlainText -Force))
                Test-OfficeDocument -LiteralPath $path -Check Security -Password $credential -Quiet
                """);
            Assert.True(Assert.IsType<bool>(Assert.Single(shell.Invoke()).BaseObject));
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            Assert.Empty(shell.Streams.Warning);
        } finally { Directory.Delete(root, true); }
    }

    private static Runspace CreateRunspace() {
        var state = InitialSessionState.CreateDefault();
        state.Commands.Add(new SessionStateCmdletEntry("Test-OfficeDocument", typeof(TestOfficeDocumentCommand), null));
        state.Commands.Add(new SessionStateCmdletEntry("New-OfficeExcelTestOptions", typeof(NewOfficeExcelTestOptionsCommand), null));
        var runspace = RunspaceFactory.CreateRunspace(state);
        runspace.Open();
        return runspace;
    }
}
