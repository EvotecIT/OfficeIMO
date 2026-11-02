using System.Management.Automation;
using System.Management.Automation.Runspaces;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Cmdlets.PowerPoint;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class PowerPointThemeWorkflowTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExplicitThemeRestylesOnlyTheSelectedMastersThenAppliesOverrides(bool allMasters) {
        using var document = CreateTwoMasters();
        string first = document.OpenXmlDocument!.PresentationPart!.SlideMasterParts.First().SlideMaster!.OuterXml;
        var selected = document.OpenXmlDocument!.PresentationPart!.SlideMasterParts.ElementAt(1);
        string beforeMaster = selected.SlideMaster!.OuterXml;
        string beforeLayout = selected.SlideLayoutParts.First().SlideLayout!.OuterXml;
        using var runspace = Open(); runspace.SessionStateProxy.SetVariable("deck", document);
        var result = Invoke(runspace, "$deck | Set-OfficePowerPointTheme -Theme Boardroom -Colors @{ Accent1 = 'magenta'; Accent2 = 'Accent3' } " +
            "-HeadingFont Cambria -BodyFont Calibri -Name 'Service Brief' -PassThru -Confirm:$false " +
            (allMasters ? "-AllMasters" : "-Master 2"));
        Assert.Same(document, Assert.Single(result).BaseObject);
        Assert.NotEqual(beforeMaster, selected.SlideMaster!.OuterXml);
        Assert.NotEqual(beforeLayout, selected.SlideLayoutParts.First().SlideLayout!.OuterXml);
        if (!allMasters) Assert.Equal(first, document.OpenXmlDocument!.PresentationPart!.SlideMasterParts.First().SlideMaster!.OuterXml);
        using var saved = PowerPointDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Equal("FF00FF", saved.GetThemeColor(PowerPointThemeColor.Accent1, 1));
        Assert.Equal(OfficeTheme.Boardroom.Palette.Accent3.ToRgbHex(), saved.GetThemeColor(PowerPointThemeColor.Accent2, 1), ignoreCase: true);
        Assert.Equal("Cambria", saved.GetThemeFonts(1).MajorLatin);
        Assert.Equal("Calibri", saved.GetThemeFonts(1).MinorLatin);
        Assert.Equal("Service Brief", PowerPointTemplate.Inspect(saved).Masters[1].ThemeName);
        if (allMasters) Assert.Equal("FF00FF", saved.GetThemeColor(PowerPointThemeColor.Accent1));
        else {
            Assert.NotEqual("FF00FF", saved.GetThemeColor(PowerPointThemeColor.Accent1));
        }
        Assert.Empty(saved.ValidateDocument());
    }

    [Fact]
    public void ColorAndFontOverridesKeepExistingLayoutGeometryAndWhatIfKeepsEveryPart() {
        using var document = CreateTwoMasters();
        using var runspace = Open(); runspace.SessionStateProxy.SetVariable("deck", document);
        string before = document.OpenXmlDocument!.PresentationPart!.SlideMasterParts.First().SlideMaster!.OuterXml;
        string theme = document.OpenXmlDocument!.PresentationPart!.SlideMasterParts.First().ThemePart!.Theme!.OuterXml;
        Assert.Empty(Invoke(runspace, "Set-OfficePowerPointTheme -Document $deck -Theme Dark -Name 'Ignored' -WhatIf"));
        Assert.Equal(before, document.OpenXmlDocument!.PresentationPart!.SlideMasterParts.First().SlideMaster!.OuterXml);
        Assert.Equal(theme, document.OpenXmlDocument!.PresentationPart!.SlideMasterParts.First().ThemePart!.Theme!.OuterXml);
        Assert.Empty(Invoke(runspace, "Set-OfficePowerPointTheme -Document $deck -Colors @{ Accent1 = '#FF00FF' } -BodyFont Calibri -BodyFontEastAsian 'MS Gothic' -Confirm:$false"));
        Assert.Equal(before, document.OpenXmlDocument!.PresentationPart!.SlideMasterParts.First().SlideMaster!.OuterXml);
        Assert.Equal("FF00FF", document.GetThemeColor(PowerPointThemeColor.Accent1));
        Assert.Equal("MS Gothic", document.GetThemeFonts().MinorEastAsian);
        Invoke(runspace, "Set-OfficePowerPointTheme -Document $deck -BodyFont Arial -ClearMissing -Confirm:$false");
        using var saved = PowerPointDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Equal("Arial", saved.GetThemeFonts().MinorLatin);
        Assert.True(string.IsNullOrEmpty(saved.GetThemeFonts().MinorEastAsian));
        Assert.Empty(saved.ValidateDocument());
    }

    [Theory]
    [InlineData("-Theme Dark -Colors @{ Accent1 = 'red'; Nonsense = 'blue' }")]
    [InlineData("-Theme Dark -Master 3")]
    [InlineData("-Colors @{ Accent1 = '#FF000080' }")]
    public void InvalidOverridesLeaveTheNativeThemeUnchanged(string parameters) {
        using var document = CreateTwoMasters();
        using var runspace = Open(); runspace.SessionStateProxy.SetVariable("deck", document);
        string before = document.OpenXmlDocument!.PresentationPart!.SlideMasterParts.First().ThemePart!.Theme!.OuterXml;
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript("Set-OfficePowerPointTheme -Document $deck " + parameters + " -Confirm:$false");
        Exception? failure = Record.Exception(() => shell.Invoke());
        Assert.True(failure != null || shell.HadErrors);
        Assert.Equal(before, document.OpenXmlDocument!.PresentationPart!.SlideMasterParts.First().ThemePart!.Theme!.OuterXml);
    }

    private static PowerPointDocument CreateTwoMasters() {
        var document = PowerPointDocument.Create();
        using var source = PowerPointDocument.Create();
        PresentationPart part = document.OpenXmlDocument!.PresentationPart!;
        SlideMasterPart added = part.AddPart(source.OpenXmlDocument!.PresentationPart!.SlideMasterParts.First());
        part.Presentation!.SlideMasterIdList!.Append(new SlideMasterId { Id = 2147483649U, RelationshipId = part.GetIdOfPart(added) });
        document.AddSlide();
        return document;
    }
    private static Runspace Open() {
        var state = InitialSessionState.CreateDefault();
        state.Commands.Add(new SessionStateCmdletEntry("Set-OfficePowerPointTheme", typeof(SetOfficePowerPointThemeCommand), null));
        var runspace = RunspaceFactory.CreateRunspace(state); runspace.Open(); return runspace;
    }
    private static System.Collections.ObjectModel.Collection<PSObject> Invoke(Runspace runspace, string script) {
        using var shell = PowerShell.Create(); shell.Runspace = runspace; shell.AddScript(script);
        var output = shell.Invoke();
        Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error.Select(error => error.ToString())));
        return output;
    }
}
