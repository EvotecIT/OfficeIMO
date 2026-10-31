using System.Management.Automation;
using System.Security;
using OfficeIMO;
using OfficeIMO.PowerPoint;

namespace PSWriteOffice.Tests;

public sealed partial class PowerPointContentWorkflowTests {
    [Fact]
    public void PresentationValidationUsesUnsavedNativeStateAndLeavesTheAssociatedFileUnchanged() {
        string path = Path.Combine(_root, "saved.pptx");
        using var document = CreateValidationDeck(); document.Save(path);
        byte[] original = File.ReadAllBytes(path);
        document.AddSlide().AddTitle("Unsaved second slide");
        using var runspace = Open(); runspace.SessionStateProxy.SetVariable("deck", document);
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript("$deck | Test-OfficePowerPoint -Options (New-OfficePowerPointInspectionOptions -Features)");
        var result = shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        var report = Assert.IsType<PowerPointInspectionReport>(Assert.Single(result).BaseObject);
        Assert.Equal(2, report.Preflight!.SlideCount); Assert.NotNull(report.Features);
        Assert.Equal(original, File.ReadAllBytes(path));
        document.AddSlide(); Assert.Equal(3, document.Slides.Count);
    }

    [Theory]
    [InlineData("Test-OfficePowerPoint -LiteralPath $path")]
    [InlineData("$file | Test-OfficePowerPoint")]
    [InlineData("Test-OfficePowerPoint -Path $pattern")]
    public void PresentationFileValidationResolvesLiteralAndWildcardInputsWithoutRetainingTheFile(string command) {
        string path = Path.Combine(_root, "review[1].pptx");
        using (var document = CreateValidationDeck()) document.Save(path);
        byte[] original = File.ReadAllBytes(path);
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("path", path);
        runspace.SessionStateProxy.SetVariable("file", new FileInfo(path));
        runspace.SessionStateProxy.SetVariable("pattern", Path.Combine(_root, "review*.pptx"));
        using var shell = PowerShell.Create(); shell.Runspace = runspace; shell.AddScript(command);
        var result = shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.Equal(1, Assert.IsType<PowerPointInspectionReport>(Assert.Single(result).BaseObject).Preflight!.SlideCount);
        Assert.Equal(original, File.ReadAllBytes(path));
        using var exclusive = File.Open(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        Assert.True(exclusive.Length > 0);
    }

    [Fact]
    public void PresentationValidationAcceptsSecurePasswordsWithoutChangingEncryptedSources() {
        string path = Path.Combine(_root, "protected.pptx");
        using (var document = CreateValidationDeck()) document.Save(path, new PowerPointSaveOptions { Password = "inspection-fixture" });
        byte[] original = File.ReadAllBytes(path);
        using var password = new SecureString();
        foreach (char character in "inspection-fixture") password.AppendChar(character);
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("path", path); runspace.SessionStateProxy.SetVariable("password", password);
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript("Test-OfficePowerPoint -LiteralPath $path -Password $password");
        var result = shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.Equal(1, Assert.IsType<PowerPointInspectionReport>(Assert.Single(result).BaseObject).Preflight!.SlideCount);
        Assert.Equal(original, File.ReadAllBytes(path));
        using var exclusive = File.Open(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
    }

    [Fact]
    public void PresentationValidationRejectsAnotherFormatWhileKeepingValidBatchResults() {
        string path = Path.Combine(_root, "review.pptx");
        string other = Path.Combine(_root, "word.docx");
        using (var document = CreateValidationDeck()) document.Save(path);
        using (var word = OfficeIMO.Word.WordDocument.Create()) { word.AddParagraph("Word source"); word.Save(other); }
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("path", path); runspace.SessionStateProxy.SetVariable("other", other);
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript("Test-OfficePowerPoint -LiteralPath $other,$path -ThrottleLimit 1 -ErrorAction Continue");
        var result = shell.Invoke();
        Assert.True(shell.HadErrors); Assert.Single(shell.Streams.Error);
        Assert.Equal(1, Assert.IsType<PowerPointInspectionReport>(Assert.Single(result).BaseObject).Preflight!.SlideCount);
    }

    [Fact]
    public void PresentationValidationAcceptsNativePolicyHashtables() {
        using var document = CreateValidationDeck();
        using var runspace = Open(); runspace.SessionStateProxy.SetVariable("deck", document);
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript("Test-OfficePowerPoint -Document $deck -Options @{ ValidatePackage = $false; InspectPreflight = $false; InspectAccessibility = $false; InspectFeatures = $true }");
        var result = shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        var report = Assert.IsType<PowerPointInspectionReport>(Assert.Single(result).BaseObject);
        Assert.Empty(report.PackageErrors); Assert.Null(report.Preflight); Assert.Null(report.Accessibility);
        Assert.NotNull(report.Features);
    }

    [Fact]
    public void PresentationValidationResolvesItsEnclosingDocumentWithoutCreatingNotes() {
        string path = Path.Combine(_root, "scoped.pptx");
        using var runspace = Open(); runspace.SessionStateProxy.SetVariable("path", path);
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript("""
            New-OfficePowerPoint -Path $path {
                PowerPointSlide -Title Review
                $global:inspection = Test-OfficePowerPoint
            }
            """);
        Assert.Empty(shell.Invoke()); Assert.False(shell.HadErrors, Errors(shell));
        object value = runspace.SessionStateProxy.GetVariable("inspection");
        var report = Assert.IsType<PowerPointInspectionReport>(value is PSObject wrapped ? wrapped.BaseObject : value);
        Assert.Equal(1, report.Preflight!.SlideCount);
        using var saved = PowerPointDocument.Load(path);
        Assert.False(Assert.Single(saved.Slides).HasSpeakerNotes);
    }

    private static PowerPointDocument CreateValidationDeck() {
        var document = PowerPointDocument.Create(new PowerPointCreateOptions { PersistenceMode = DocumentPersistenceMode.Explicit });
        document.AddSlide().AddTitle("Saved review");
        return document;
    }
}
