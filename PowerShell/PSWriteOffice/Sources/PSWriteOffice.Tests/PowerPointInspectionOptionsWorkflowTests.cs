using System.Management.Automation;
using System.Management.Automation.Runspaces;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Cmdlets.PowerPoint;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class PowerPointInspectionOptionsWorkflowTests {
    [Fact]
    public void DefaultPolicyRunsTheDefaultNativeReportsOnAnUnsavedPresentation() {
        using var document = Create();
        var report = document.Inspect(Options("New-OfficePowerPointInspectionOptions"));
        Assert.Empty(report.PackageErrors);
        Assert.NotNull(report.Preflight); Assert.NotNull(report.Accessibility);
        Assert.Null(report.Features); Assert.Null(report.ReviewComments);
        Assert.Null(report.Animations); Assert.Null(report.Signatures); Assert.Null(report.Visuals);
        Assert.False(document.Slides[0].HasSpeakerNotes);
    }

    [Fact]
    public void SelectedPolicyRunsOptInReportsAndProducesNativeVisualEvidence() {
        using var document = Create();
        var policy = Options("New-OfficePowerPointInspectionOptions -NoPackageValidation -NoPreflight -NoAccessibility -Features -ReviewComments -Animations -Signatures -Visuals -FileFormatVersion Office2021");
        var report = document.Inspect(policy);
        Assert.Empty(report.PackageErrors); Assert.Null(report.Preflight); Assert.Null(report.Accessibility);
        Assert.NotNull(report.Features); Assert.NotNull(report.ReviewComments);
        Assert.NotNull(report.Animations); Assert.NotNull(report.Signatures);
        var slide = Assert.Single(Assert.IsType<PowerPointVisualProofReport>(report.Visuals).Slides);
        Assert.True(slide.Png.ByteLength > 0); Assert.True(slide.Svg.ByteLength > 0);
        Assert.Equal(OfficeIMO.OfficeOpenXmlFileFormatVersion.Office2021, policy.FileFormatVersion);
        Assert.False(document.Slides[0].HasSpeakerNotes);
    }

    [Fact]
    public void ExplicitFalseSwitchesPreserveNativeDefaultsWithoutChangingAnotherPolicy() {
        using var document = Create();
        var selected = Options("New-OfficePowerPointInspectionOptions -NoPreflight -Features");
        var defaults = Options("New-OfficePowerPointInspectionOptions -NoPackageValidation:$false -NoPreflight:$false -NoAccessibility:$false -Features:$false -Visuals:$false");
        var defaultReport = document.Inspect(defaults);
        Assert.NotNull(defaultReport.Preflight); Assert.NotNull(defaultReport.Accessibility);
        Assert.Null(defaultReport.Features); Assert.Null(defaultReport.Visuals);
        var selectedReport = document.Inspect(selected);
        Assert.Null(selectedReport.Preflight); Assert.NotNull(selectedReport.Features);
    }

    [Fact]
    public void DetailedDocumentInfoConsumesTheTypedPolicyWithoutReopeningTheLiveDeck() {
        using var document = Create();
        using var runspace = Open(); runspace.SessionStateProxy.SetVariable("deck", document);
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript("""
            $policy = New-OfficePowerPointInspectionOptions -NoPackageValidation -NoPreflight -Features
            (Get-OfficeDocumentInfo -Document $deck -Detailed -Options $policy).Details
            """);
        var result = shell.Invoke();
        Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error.Select(error => error.ToString())));
        var report = Assert.IsType<PowerPointInspectionReport>(Assert.Single(result).BaseObject);
        Assert.Null(report.Preflight); Assert.NotNull(report.Features); Assert.NotNull(report.Accessibility);
        Assert.False(document.Slides[0].HasSpeakerNotes);
        Assert.Equal("Review", document.Slides[0].GetPlaceholder(PowerPointPlaceholderType.Title)!.Text);
    }

    private static PowerPointPresentation Create() {
        var document = PowerPointPresentation.Create();
        document.AddSlide().AddTitle("Review");
        return document;
    }

    private static PowerPointInspectionOptions Options(string script) {
        using var runspace = Open();
        using var shell = PowerShell.Create(); shell.Runspace = runspace; shell.AddScript(script);
        var result = shell.Invoke();
        Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error.Select(error => error.ToString())));
        return Assert.IsType<PowerPointInspectionOptions>(Assert.Single(result).BaseObject);
    }

    private static Runspace Open() {
        var state = InitialSessionState.CreateDefault();
        state.Commands.Add(new SessionStateCmdletEntry("New-OfficePowerPointInspectionOptions", typeof(NewOfficePowerPointInspectionOptionsCommand), null));
        state.Commands.Add(new SessionStateCmdletEntry("Get-OfficeDocumentInfo", typeof(PSWriteOffice.Cmdlets.Reader.GetOfficeDocumentInfoCommand), null));
        var runspace = RunspaceFactory.CreateRunspace(state); runspace.Open(); return runspace;
    }
}
