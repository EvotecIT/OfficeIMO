using System.Management.Automation;
using System.Management.Automation.Runspaces;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Cmdlets.PowerPoint;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class PowerPointSlideSizeWorkflowTests {
    [Theory]
    [InlineData("-Width '25.4cm' -Height '14cm'", 9144000L, 5040000L)]
    [InlineData("-Width 720 -Height 405", 9144000L, 5143500L)]
    public void PhysicalLengthsAndBarePointsSurviveNativeSaveAndReopen(string parameters, long width, long height) {
        using var document = PowerPointDocument.Create();
        using var runspace = Open(); runspace.SessionStateProxy.SetVariable("deck", document);
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript("$deck | Set-OfficePowerPointSlideSize " + parameters + " -PassThru");
        Assert.Same(document.SlideSize, Assert.Single(shell.Invoke()).BaseObject);
        Assert.False(shell.HadErrors, Errors(shell));
        document.AddSlide().AddTitle("Physical sizing");
        using var input = new MemoryStream(document.ToBytes());
        using var reopened = PowerPointDocument.Load(input);
        Assert.Equal(width, reopened.SlideSize.Width.ToEmus());
        Assert.Equal(height, reopened.SlideSize.Height.ToEmus());
        Assert.Empty(reopened.ValidateDocument());
    }

    [Fact]
    public void PortraitPresetUsesTheNativeSizeAndIsQuietByDefault() {
        using var document = PowerPointDocument.Create();
        using var runspace = Open(); runspace.SessionStateProxy.SetVariable("deck", document);
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript("Set-OfficePowerPointSlideSize -Document $deck -Preset Screen4x3 -Portrait");
        Assert.Empty(shell.Invoke());
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.Equal(6858000L, document.SlideSize.Width.ToEmus());
        Assert.Equal(9144000L, document.SlideSize.Height.ToEmus());
    }

    [Fact]
    public void InvalidHeightReportsAnArgumentErrorAndLeavesBothDimensionsUnchanged() {
        using var document = PowerPointDocument.Create();
        long width = document.SlideSize.Width.ToEmus(), height = document.SlideSize.Height.ToEmus();
        using var runspace = Open(); runspace.SessionStateProxy.SetVariable("deck", document);
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript("Set-OfficePowerPointSlideSize -Document $deck -Width '1in' -Height '2147483648emu'");
        Exception? failure = Record.Exception(() => shell.Invoke());
        Assert.True(failure != null || shell.HadErrors);
        Assert.Equal(ErrorCategory.InvalidArgument, Assert.Single(shell.Streams.Error).CategoryInfo.Category);
        Assert.Equal(width, document.SlideSize.Width.ToEmus());
        Assert.Equal(height, document.SlideSize.Height.ToEmus());
    }

    private static Runspace Open() {
        var state = InitialSessionState.CreateDefault();
        state.Commands.Add(new SessionStateCmdletEntry("Set-OfficePowerPointSlideSize", typeof(SetOfficePowerPointSlideSizeCommand), null));
        var runspace = RunspaceFactory.CreateRunspace(state); runspace.Open(); return runspace;
    }
    private static string Errors(PowerShell shell) => string.Join(Environment.NewLine, shell.Streams.Error.Select(error => error.ToString()));
}
