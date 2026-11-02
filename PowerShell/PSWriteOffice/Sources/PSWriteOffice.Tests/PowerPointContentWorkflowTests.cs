using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Reflection;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Cmdlets.PowerPoint;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed partial class PowerPointContentWorkflowTests : IDisposable {
    private readonly string _root = Path.Combine(Path.GetTempPath(), "officeimo-ppt-content-" + Guid.NewGuid().ToString("N"));
    public PowerPointContentWorkflowTests() => Directory.CreateDirectory(_root);
    public void Dispose() => Directory.Delete(_root, true);

    [Fact]
    public void SpelledOutContentCommandsProduceEditableDesignedSlides() {
        string path = Path.Combine(_root, "service-review.pptx");
        using var runspace = Open();
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        runspace.SessionStateProxy.SetVariable("target", path);
        shell.AddScript("""
            New-OfficePowerPoint -Path $target -Theme Modern -PassThru {
                $layout = @(Get-OfficePowerPointLayout)[0]
                PowerPointSlide -Title 'Service review' -Subtitle 'October 2026' -TemplateLayout $layout
                PowerPointSlide -Title 'Current signals' {
                    PowerPointCard 'Availability' -Text 'All regions healthy'
                    PowerPointCard 'Capacity' -Text 'Storage at 71%'
                }
                PowerPointSlide -Title 'Quarterly comparison' {
                    PowerPointColumn -Title 'Current' {
                        PowerPointParagraph 'Support priorities'
                        PowerPointList 'Availability', 'Response time' -Numbered
                    }
                    PowerPointColumn -Title 'Next' {
                        PowerPointList 'Resilience', 'Capacity'
                    }
                }
                PowerPointSlide -Title 'Health' {
                    PowerPointKpi '99.95%' 'Availability' -Change '+0.02 pt'
                    PowerPointKpi '182 ms' 'Median latency' -Color Warning
                }
                PowerPointSlide -Title 'Operating rhythm' {
                    PowerPointStep 'Collect' 'Gather health signals'
                    PowerPointStep 'Review' 'Confirm decisions with owners'
                    PowerPointStep 'Act' 'Track actions to closure'
                }
                PowerPointSlide { PowerPointQuote 'A reliable service starts with clear ownership.' 'Operations team' }
            }
            """);
        var result = shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.Equal(path, Assert.IsType<FileInfo>(Assert.Single(result).BaseObject).FullName);
        using var saved = PowerPointDocument.Load(path);
        Assert.Equal(6, saved.Slides.Count);
        Assert.Empty(saved.ValidateDocument());
        string[] expected = { "Service review", "Support priorities", "Capacity", "All regions healthy", "99.95%", "Track actions to closure", "Operations team" };
        foreach (string text in expected)
            Assert.Contains(saved.Slides.SelectMany(slide => slide.TextBoxes), box => box.Text.Contains(text, StringComparison.Ordinal));
    }

    [Fact]
    public void AFailedSlideBlockDoesNotAppendAnEmptySlideOrLeakItsCollector() {
        using var document = PowerPointDocument.Create();
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("deck", document);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript("Add-OfficePowerPointSlide -Document $deck -Title 'Not appended' { throw 'content-failed' }");
        Exception? failure = Record.Exception(() => shell.Invoke());
        Assert.True(failure != null || shell.HadErrors);
        Assert.Empty(document.Slides);
        shell.Commands.Clear();
        shell.Streams.Error.Clear();
        shell.AddScript("Add-OfficePowerPointSlide -Document $deck -Title 'Recovered' -PassThru { PowerPointCard 'Ready' -Text 'Clean scope' }");
        Assert.IsType<PowerPointSlide>(Assert.Single(shell.Invoke()).BaseObject);
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.Single(document.Slides);
    }

    [Fact]
    public void NestedNewUsesItsOwnDocumentAndRestoresTheOuterSlideCollector() {
        string outer = Path.Combine(_root, "outer.pptx"), inner = Path.Combine(_root, "inner.pptx");
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("outer", outer);
        runspace.SessionStateProxy.SetVariable("inner", inner);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript("""
            New-OfficePowerPoint -Path $outer {
                PowerPointSlide -Title 'Outer' {
                    PowerPointNotes 'Outer talking points'
                    PowerPointCard 'Before' -Text 'Outer content'
                    New-OfficePowerPoint -Path $inner {
                        PowerPointSlide -Title 'Inner' {
                            PowerPointNotes 'Inner talking points'
                            PowerPointCard 'Private' -Text 'Inner content'
                        }
                    }
                    PowerPointCard 'After' -Text 'Outer scope restored'
                }
            }
            """);
        shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        using var outerDeck = PowerPointDocument.Load(outer);
        using var innerDeck = PowerPointDocument.Load(inner);
        Assert.Single(outerDeck.Slides);
        Assert.Single(innerDeck.Slides);
        string outerText = string.Join("\n", outerDeck.Slides[0].TextBoxes.Select(box => box.Text));
        string innerText = string.Join("\n", innerDeck.Slides[0].TextBoxes.Select(box => box.Text));
        Assert.Contains("Before", outerText);
        Assert.Contains("After", outerText);
        Assert.DoesNotContain("Private", outerText);
        Assert.Contains("Private", innerText);
        Assert.DoesNotContain("Before", innerText);
        Assert.Equal("Outer talking points", outerDeck.Slides[0].Notes.Text);
        Assert.Equal("Inner talking points", innerDeck.Slides[0].Notes.Text);
        Assert.DoesNotContain("talking points", outerText);
        Assert.DoesNotContain("talking points", innerText);
    }

    [Fact]
    public void ColumnNotesBelongToTheSlideAndNativeNoteEditsReturnTheSelectedSlide() {
        string path = Path.Combine(_root, "notes.pptx");
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("target", path);
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript("""
            New-OfficePowerPoint -Path $target {
                PowerPointSlide -Title 'Notes' {
                    PowerPointColumn {
                        PowerPointParagraph 'Visible body'
                        PowerPointNotes 'Private talking points'
                    }
                }
                PowerPointSlide -Title 'Unchanged' { PowerPointNotes 'Other notes' }
            }
            """);
        Assert.Empty(shell.Invoke());
        Assert.False(shell.HadErrors, Errors(shell));
        using var document = PowerPointDocument.Load(path);
        Assert.Equal("Private talking points", document.Slides[0].Notes.Text);
        Assert.DoesNotContain(document.Slides[0].TextBoxes, box => box.Text.Contains("Private talking points", StringComparison.Ordinal));
        runspace.SessionStateProxy.SetVariable("deck", document);
        foreach (string selection in new[] {
            "Set-OfficePowerPointNotes -Document $deck -Slide 1",
            "Set-OfficePowerPointNotes -Slide $deck.Slides[0]",
            "$deck.Slides[0] | Set-OfficePowerPointNotes"
        }) {
            shell.Commands.Clear();
            shell.AddScript(selection + " -Text '' -PassThru");
            Assert.Same(document.Slides[0], Assert.Single(shell.Invoke()).BaseObject);
            Assert.False(shell.HadErrors, Errors(shell));
            Assert.Equal(string.Empty, document.Slides[0].Notes.Text);
        }
        Assert.Equal("Other notes", document.Slides[1].Notes.Text);
        using var input = new MemoryStream(document.ToBytes());
        using var reopened = PowerPointDocument.Load(input);
        Assert.Equal(string.Empty, reopened.Slides[0].Notes.Text);
        Assert.Equal("Other notes", reopened.Slides[1].Notes.Text);
        Assert.Empty(reopened.ValidateDocument());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PipelineTableProjectsSelectedPropertiesAndContinuesWithoutLosingRows(bool noHeader) {
        string path = Path.Combine(_root, "table.pptx");
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("target", path);
        runspace.SessionStateProxy.SetVariable("noHeader", noHeader);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript("""
            New-OfficePowerPoint -Path $target {
                PowerPointSlide -Title 'Services' {
                    1..7 | ForEach-Object { [pscustomobject]@{ Service = "Service $_"; Count = $_; Hidden = 'private' } } |
                        PowerPointTable -Property Service, Count -MaxRows 3 -NoHeader:$noHeader
                }
            }
            """);
        shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        using var document = PowerPointDocument.Load(path);
        Assert.Equal(3, document.Slides.Count);
        var values = new List<string>();
        foreach (var table in document.Slides.SelectMany(slide => slide.Tables)) {
            Assert.Equal(2, table.Columns);
            if (!noHeader) Assert.Equal("Service", table.GetCell(0, 0).Text);
            for (int row = noHeader ? 0 : 1; row < table.Rows; row++) values.Add(table.GetCell(row, 0).Text);
        }
        Assert.Equal(Enumerable.Range(1, 7).Select(index => "Service " + index), values);
        Assert.Empty(document.ValidateDocument());
    }

    [Theory]
    [InlineData("-Change '14 ms lower' -Trend Down -GoodDirection Down", "▼ 14 ms lower", "success")]
    [InlineData("-Delta -0.14 -DeltaUnit '%' -GoodDirection Down", "▼ -14%", "success")]
    [InlineData("-Delta 14 -DeltaUnit ms -GoodDirection Down -Color '#123456'", "▲ +14 ms", "123456")]
    [InlineData("-Delta 0 -DeltaUnit '%'", "0%", null)]
    public void KpiChangeBindingsPreserveArrowTextAndSemanticColor(string parameters, string expected, string? color) {
        string path = Path.Combine(_root, "metric.pptx");
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("target", path);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript("New-OfficePowerPoint -Path $target { PowerPointSlide -Title 'Latency' { PowerPointKpi '182 ms' 'Median response' " + parameters + " } }");
        shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        using var document = PowerPointDocument.Load(path);
        var change = Assert.Single(Assert.Single(document.Slides).TextBoxes, box => box.Text == expected);
        Assert.Equal(expected, change.Text);
        if (color != null) {
            string wanted = color == "success" ? OfficeIMO.OfficeTheme.Modern.Palette.Success.ToRgbHex() : color;
            Assert.Equal(wanted, Assert.Single(change.Paragraphs.SelectMany(paragraph => paragraph.Runs)).Color, ignoreCase: true);
        }
        Assert.Empty(document.ValidateDocument());
    }

    [Theory]
    [InlineData("position")]
    [InlineData("name")]
    [InlineData("type")]
    [InlineData("object")]
    public void TemplateLayoutSelectsTheSameNativeLayoutByPositionNameTypeOrObject(string selectionKind) {
        using var document = PowerPointDocument.Create();
        var layout = document.GetSlideLayouts(0).First(entry => entry.Type.HasValue);
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("deck", document);
        object selection = selectionKind switch {
            "position" => layout.LayoutIndex + 1,
            "name" => layout.Name,
            "type" => layout.Type!.Value,
            _ => layout
        };
        runspace.SessionStateProxy.SetVariable("selection", selection);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        string choose = selectionKind == "object" ? "$selection = @(Get-OfficePowerPointLayout -Document $deck -Master 1 | Where-Object { $null -ne $_.Type })[0]; " : string.Empty;
        shell.AddScript(choose + "Add-OfficePowerPointSlide -Document $deck -TemplateLayout $selection -Master 1 -Title 'Template evidence' -PassThru { PowerPointList 'Preserve the chosen layout' }");
        var result = shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        var slide = Assert.IsType<PowerPointSlide>(Assert.Single(result).BaseObject);
        Assert.Equal(layout.LayoutIndex, slide.LayoutIndex);
        Assert.Contains(slide.TextBoxes, box => box.Text.Contains("Preserve the chosen layout", StringComparison.Ordinal));
        Assert.Empty(document.ValidateDocument());
    }

    [Fact]
    public void MasterWithoutTemplateLayoutIsRejectedBeforeChangingTheDeck() {
        using var document = PowerPointDocument.Create();
        var original = document.AddSlide();
        original.AddTitle("Retained");
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("deck", document);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript("PowerPointSlide -Document $deck -Master 1 -Title 'Must not be created'");
        Exception? error = Record.Exception(() => shell.Invoke());
        Assert.True(error != null || shell.HadErrors);
        var failure = error?.GetBaseException() ?? Assert.Single(shell.Streams.Error).Exception.GetBaseException();
        Assert.IsType<ArgumentException>(failure);
        Assert.Same(original, Assert.Single(document.Slides));
        Assert.Equal("Retained", original.GetPlaceholder(PowerPointPlaceholderType.Title)!.Text);
    }

    [Fact]
    public void MetadataReadersUseTheSharedNewDocumentScope() {
        string path = Path.Combine(_root, "metadata.pptx");
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("target", path);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript("""
            New-OfficePowerPoint -Path $target -Theme Modern {
                PowerPointSlide -Title 'Metadata review'
                $global:theme = Get-OfficePowerPointTheme -Master 1
                $global:sections = @(Get-OfficePowerPointSection)
            }
            """);
        shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        object value = runspace.SessionStateProxy.GetVariable("theme");
        var theme = Assert.IsType<PSWriteOffice.Services.PowerPoint.PowerPointThemeInfo>(value is PSObject wrapped ? wrapped.BaseObject : value);
        Assert.NotEmpty(theme.Colors);
        using var saved = PowerPointDocument.Load(path);
        Assert.Single(saved.Slides);
        Assert.Equal(saved.ThemeName, theme.ThemeName);
    }

    [Fact]
    public void SectionRenameAndListingUseDocumentArgumentsAndPipelineInput() {
        using var document = PowerPointDocument.Create();
        document.AddSlide();
        document.AddSection("Review", 0);
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("deck", document);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript("Rename-OfficePowerPointSection -Document $deck -Name Review -NewName Accepted -PassThru");
        var renamed = shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.Equal("Accepted", Assert.IsType<PowerPointSectionInfo>(Assert.Single(renamed).BaseObject).Name);
        shell.Commands.Clear();
        shell.AddScript("$deck | Get-OfficePowerPointSection -Name Accepted");
        var found = shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.Equal("Accepted", Assert.IsType<PowerPointSectionInfo>(Assert.Single(found).BaseObject).Name);
    }

    [Fact]
    public void RenamingToAnExistingSectionNameReturnsTheSectionThatWasChanged() {
        using var document = PowerPointDocument.Create();
        document.AddSlide();
        document.AddSlide();
        document.AddSection("Accepted", 0);
        string id = document.AddSection("Review", 1).Id;
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("deck", document);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript("Rename-OfficePowerPointSection -Document $deck -Name Review -NewName Accepted -PassThru");
        var result = shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.Equal(id, Assert.IsType<PowerPointSectionInfo>(Assert.Single(result).BaseObject).Id);
    }

    private static Runspace Open() {
        var state = InitialSessionState.CreateDefault();
        foreach (Type type in typeof(NewOfficePowerPointCommand).Assembly.GetTypes()) {
            var command = type.GetCustomAttribute<CmdletAttribute>();
            if (command == null || !command.NounName.StartsWith("OfficePowerPoint", StringComparison.Ordinal)) continue;
            string name = command.VerbName + "-" + command.NounName;
            state.Commands.Add(new SessionStateCmdletEntry(name, type, null));
            foreach (string alias in type.GetCustomAttribute<AliasAttribute>()?.AliasNames ?? Array.Empty<string>())
                state.Commands.Add(new SessionStateAliasEntry(alias, name));
        }
        var runspace = RunspaceFactory.CreateRunspace(state);
        runspace.Open();
        return runspace;
    }

    private static string Errors(PowerShell shell) => string.Join(Environment.NewLine, shell.Streams.Error.Select(error => error.ToString()));
}
