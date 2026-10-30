using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Reflection;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Cmdlets.PowerPoint;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed partial class PowerPointNativeOutputWorkflowTests : IDisposable {
    private readonly string _root = Path.Combine(Path.GetTempPath(), "officeimo-ppt-native-" + Guid.NewGuid().ToString("N"));
    public PowerPointNativeOutputWorkflowTests() => Directory.CreateDirectory(_root);
    public void Dispose() => Directory.Delete(_root, true);

    [Fact]
    public void RepeatedParagraphsReturnDistinctLiveNodesAndPreserveRichRunsAndNotes() {
        using var deck = PowerPointPresentation.Create();
        using var runspace = Open(deck);
        Invoke(runspace, """
            $script:physical = PowerPointSlide -Document $deck -Title 'Status' -PassThru {
                $script:first = PowerPointParagraph 'Repeated' -PassThru
                $script:firstRun = $first.AddRun(' linked')
                $firstRun.Bold = $true
                $firstRun.SetHyperlink('https://contoso.com/before')
                $script:second = PowerPointParagraph 'Repeated' -PassThru
                $script:noteSlide = PowerPointNotes 'Original notes' -PassThru
                $noteSlide.Notes.Text = 'Edited during collection'
                $first.Text = 'First edited during collection'
                $script:richRun = $first.AddRun(' linked')
                $richRun.Italic = $true
                $richRun.SetHyperlink('https://contoso.com/live')
            }
            $second.Text = 'Second edited after layout'
            $richRun.Text = ' live link'
            $noteSlide.Notes.Text = 'Edited after layout'
            """);
        var first = Get<PowerPointParagraph>(runspace, "first");
        var second = Get<PowerPointParagraph>(runspace, "second");
        var richRun = Get<PowerPointTextRun>(runspace, "richRun");
        Assert.NotSame(Node(first), Node(second));
        AssertBacked(deck, first);
        AssertBacked(deck, second);
        Assert.Same(deck.Slides[0], Get<PowerPointSlide>(runspace, "noteSlide"));
        Assert.Same(deck.Slides[0], Get<PowerPointSlide>(runspace, "physical"));
        Assert.Same(Node(richRun), Node(first.Runs.Last()));
        string path = Save(deck, "repeated.pptx");
        using var saved = PowerPointPresentation.Load(path);
        var paragraphs = saved.Slides.SelectMany(s => s.TextBoxes).SelectMany(b => b.Paragraphs).ToList();
        Assert.Contains(paragraphs, p => p.Text == "Second edited after layout");
        var stored = Assert.Single(paragraphs, p => p.Text == "First edited during collection live link");
        Assert.True(stored.Runs.Last().Italic);
        Assert.Equal("https://contoso.com/live", stored.Runs.Last().Hyperlink!.AbsoluteUri);
        Assert.Equal("Edited after layout", saved.Slides[0].GetSpeakerNotesText());
        Assert.Empty(saved.ValidateDocument());
    }

    [Theory]
    [InlineData("Auto")]
    [InlineData("Quote")]
    [InlineData("Closing")]
    public void RichParagraphCommandsKeepNativeBackingInBodyQuoteAndClosing(string layout) {
        using var deck = PowerPointPresentation.Create();
        using var runspace = Open(deck);
        runspace.SessionStateProxy.SetVariable("layout", layout);
        Invoke(runspace, """
            PowerPointSlide -Document $deck -Title 'Decision' -Layout $layout {
                $script:paragraph = PowerPointParagraph -PassThru {
                    PowerPointText 'Status: ' -Bold
                    $script:run = PowerPointText 'Healthy' -Italic -Underline Double -Color '#127733' -Link 'https://contoso.com/status' -PassThru
                }
                $run.Text = 'Ready'
            }
            $paragraph.AddText(' today') | Out-Null
            $run.SetHyperlink('https://contoso.com/updated')
            """);
        var paragraph = Get<PowerPointParagraph>(runspace, "paragraph");
        AssertBacked(deck, paragraph);
        string path = Save(deck, "rich-" + layout + ".pptx");
        using var saved = PowerPointPresentation.Load(path);
        var stored = Assert.Single(saved.Slides.SelectMany(s => s.TextBoxes).SelectMany(b => b.Paragraphs),
            p => p.Text == "Status: Ready today");
        Assert.True(stored.Runs.Single(r => r.Text == "Status: ").Bold);
        var linked = stored.Runs.Single(r => r.Text == "Ready");
        Assert.True(linked.Italic);
        Assert.Equal(PowerPointUnderlineStyle.Double, linked.UnderlineStyle);
        Assert.Equal("127733", linked.Color);
        Assert.Equal("https://contoso.com/updated", linked.Hyperlink!.AbsoluteUri);
        Assert.Empty(saved.ValidateDocument());
    }

    [Fact]
    public void ComparisonColumnsAndTheirNotesShareOnlyTheSlideNativeOwner() {
        using var deck = PowerPointPresentation.Create();
        using var runspace = Open(deck);
        Invoke(runspace, """
            PowerPointSlide -Document $deck -Title 'Comparison' {
                PowerPointColumn -Title 'Before' {
                    $script:lede = PowerPointParagraph 'Repeated' -PassThru
                    $script:left = PowerPointList 'Repeated', 'Repeated' -PassThru
                    $script:notes = PowerPointNotes 'Column notes' -PassThru
                    $left.Paragraphs[1].Text = 'Changed left during collection'
                }
                PowerPointColumn -Title 'After' {
                    $script:right = PowerPointList -PassThru {
                        $script:item = PowerPointListItem -PassThru {
                            PowerPointText 'Repeated' -Bold
                            $script:link = PowerPointText ' link' -Link 'https://contoso.com/right' -PassThru
                        }
                    }
                    $item.SetBullet([char]'*')
                }
            }
            $lede.Text = 'Changed lede after layout'
            $item.AddText(' after layout') | Out-Null
            $notes.Notes.Text = 'Changed column notes'
            """);
        var left = Get<PowerPointTextBox>(runspace, "left");
        var right = Get<PowerPointTextBox>(runspace, "right");
        Assert.Contains(deck.Slides[0].TextBoxes, b => ReferenceEquals(b, left));
        Assert.Contains(deck.Slides[0].TextBoxes, b => ReferenceEquals(b, right));
        AssertBacked(deck, Get<PowerPointParagraph>(runspace, "lede"));
        AssertBacked(deck, Get<PowerPointParagraph>(runspace, "item"));
        using var saved = PowerPointPresentation.Load(Save(deck, "columns.pptx"));
        Assert.Equal("Changed column notes", saved.Slides[0].GetSpeakerNotesText());
        string text = string.Join("\n", saved.Slides[0].TextBoxes.Select(b => b.Text));
        Assert.Contains("Changed lede after layout", text);
        Assert.Contains("Changed left during collection", text);
        Assert.Contains("Repeated link after layout", text);
        var itemNode = (DocumentFormat.OpenXml.Drawing.Paragraph)Node(Get<PowerPointParagraph>(runspace, "item"));
        Assert.Equal("*", itemNode.ParagraphProperties!.GetFirstChild<DocumentFormat.OpenXml.Drawing.CharacterBullet>()!.Char!.Value);
        Assert.Null(itemNode.ParagraphProperties.GetFirstChild<DocumentFormat.OpenXml.Drawing.AutoNumberedBullet>());
        Assert.Null(itemNode.ParagraphProperties.GetFirstChild<DocumentFormat.OpenXml.Drawing.NoBullet>());
        Assert.Empty(saved.ValidateDocument());
    }

    [Fact]
    public void EmptyListKeepsTheArrangedNativeFrameForLaterEdits() {
        using var deck = PowerPointPresentation.Create();
        using var runspace = Open(deck);
        Invoke(runspace, """
            PowerPointSlide -Document $deck -Title 'Pending actions' {
                $script:list = PowerPointList {} -PassThru
            }
            $list.AddParagraph('Added through the captured frame') | Out-Null
            """);
        var list = Get<PowerPointTextBox>(runspace, "list");
        Assert.Contains(deck.Slides[0].TextBoxes, b => ReferenceEquals(b, list));
        Assert.NotNull(list.TextAutoFit);
        Assert.Equal(PowerPointTextAutoFit.None, list.TextAutoFit);
        using var saved = PowerPointPresentation.Load(Save(deck, "empty-list.pptx"));
        Assert.Contains(saved.Slides[0].TextBoxes, b => b.Text.Contains("Added through the captured frame", StringComparison.Ordinal));
        Assert.Empty(saved.ValidateDocument());
    }

    [Fact]
    public void NumberedListContinuationKeepsTheFirstFrameAndAllCapturedParagraphsAndLinksLive() {
        using var deck = PowerPointPresentation.Create();
        using var runspace = Open(deck);
        Invoke(runspace, """
            PowerPointSlide -Document $deck -Title 'Many actions' {
                $script:list = PowerPointList -Numbered -PassThru {
                    foreach ($index in 1..42) {
                        PowerPointListItem {
                            $run = PowerPointText 'Repeated action' -PassThru
                            if ($index -eq 42) {
                                $script:lastRun = $run
                                $run.SetHyperlink('https://contoso.com/before-continuation')
                            }
                        }
                    }
                }
                $script:last = $list.Paragraphs[41]
            }
            $list.Paragraphs[0].Text = 'First action edited after layout'
            $lastRun.Text = 'Last action edited after layout'
            $lastRun.SetHyperlink('https://contoso.com/after-continuation')
            $last.AddText(' extra') | Out-Null
            """);
        Assert.True(deck.Slides.Count > 1, "The fixture must exercise physical continuation.");
        var list = Get<PowerPointTextBox>(runspace, "list");
        Assert.Contains(deck.Slides[0].TextBoxes, b => ReferenceEquals(b, list));
        var last = Get<PowerPointParagraph>(runspace, "last");
        AssertBacked(deck, last);
        Assert.Contains(deck.Slides.Skip(1).SelectMany(s => s.TextBoxes).SelectMany(b => b.Paragraphs), p => ReferenceEquals(Node(p), Node(last)));
        Assert.Equal("https://contoso.com/after-continuation", Get<PowerPointTextRun>(runspace, "lastRun").Hyperlink!.AbsoluteUri);
        using var saved = PowerPointPresentation.Load(Save(deck, "continued-list.pptx"));
        var items = saved.Slides.SelectMany(s => s.TextBoxes).SelectMany(b => b.Paragraphs)
            .Where(p => ((DocumentFormat.OpenXml.Drawing.Paragraph)Node(p)).ParagraphProperties?
                .GetFirstChild<DocumentFormat.OpenXml.Drawing.AutoNumberedBullet>() != null).ToList();
        var expected = new[] { "First action edited after layout" }.Concat(Enumerable.Repeat("Repeated action", 40))
            .Concat(new[] { "Last action edited after layout extra" });
        Assert.Equal(expected, items.Select(p => p.Text));
        Assert.Equal(Enumerable.Range(1, 42), items.Select(p =>
            ((DocumentFormat.OpenXml.Drawing.Paragraph)Node(p)).ParagraphProperties!
                .GetFirstChild<DocumentFormat.OpenXml.Drawing.AutoNumberedBullet>()!.StartAt!.Value));
        Assert.Equal("https://contoso.com/after-continuation", items.Last().Runs.First().Hyperlink!.AbsoluteUri);
        Assert.DoesNotContain(saved.OpenXmlDocument.PresentationPart!.SlideParts.SelectMany(p => p.HyperlinkRelationships),
            link => link.Uri.AbsoluteUri == "https://contoso.com/before-continuation");
        Assert.Empty(saved.ValidateDocument());
    }

    [Fact]
    public void SentenceContinuationPreservesCapturedRunsThatStartOnLaterSlides() {
        using var deck = PowerPointPresentation.Create();
        using var runspace = Open(deck);
        string repeated = string.Join(" ", Enumerable.Repeat("Repeated sentence with enough detail to fill the slide and require continuation.", 36)) + " ";
        runspace.SessionStateProxy.SetVariable("repeated", repeated);
        Invoke(runspace, """
            PowerPointSlide -Document $deck -Title 'Long narrative' {
                $script:paragraph = PowerPointParagraph -PassThru {
                    PowerPointText $repeated
                    $script:tail = PowerPointText 'Final linked sentence.' -Bold -Link 'https://contoso.com/before' -PassThru
                }
            }
            $tail.Text = 'Final edited sentence.'
            """);
        Assert.True(deck.Slides.Count > 1, "The fixture must split a native paragraph between sentences.");
        AssertBacked(deck, Get<PowerPointParagraph>(runspace, "paragraph"));
        var tail = Get<PowerPointTextRun>(runspace, "tail");
        Assert.Contains(deck.Slides.Skip(1).SelectMany(s => s.TextBoxes).SelectMany(b => b.Paragraphs).SelectMany(p => p.Runs),
            r => ReferenceEquals(Node(r), Node(tail)));
        using var saved = PowerPointPresentation.Load(Save(deck, "continued-paragraph.pptx"));
        var fragments = saved.Slides.SelectMany(s => s.TextBoxes).SelectMany(b => b.Paragraphs)
            .Where(p => p.Text.StartsWith("Repeated sentence", StringComparison.Ordinal) || p.Text.StartsWith("Final edited", StringComparison.Ordinal)).ToList();
        Assert.Equal(repeated + "Final edited sentence.", string.Concat(fragments.Select(p => p.Text)));
        var storedTail = fragments.SelectMany(p => p.Runs).Single(r => r.Text == "Final edited sentence.");
        Assert.True(storedTail.Bold);
        Assert.Equal("https://contoso.com/before", storedTail.Hyperlink!.AbsoluteUri);
        Assert.Empty(saved.ValidateDocument());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NativeSlideAndRejectedColumnRollbackRestoreScopeAndExistingContent(bool renderFailure) {
        using var deck = PowerPointPresentation.Create();
        var existing = deck.AddSlide();
        var untouched = existing.AddTextBox("Existing content");
        using var runspace = Open(deck);
        string[] initialNotes = NotesEntries(Save(deck, "before-failed-slide.pptx"));
        runspace.SessionStateProxy.SetVariable("renderFailure", renderFailure);
        runspace.SessionStateProxy.SetVariable("missingImage", Path.Combine(_root, "missing.png"));
        using (var shell = PowerShell.Create()) {
            shell.Runspace = runspace;
            shell.AddScript("""
                PowerPointSlide -Document $deck -Title 'Rejected' {
                    $script:rejected = PowerPointParagraph 'Rejected native body' -PassThru
                    PowerPointNotes 'Rejected notes' -PassThru | Out-Null
                    if ($renderFailure) { PowerPointImage $missingImage } else { throw 'slide-block-failed' }
                }
                """);
            Exception? failure = Record.Exception(() => shell.Invoke());
            Assert.True(failure != null || shell.HadErrors);
        }
        Assert.Same(existing, Assert.Single(deck.Slides));
        Assert.Same(untouched, Assert.Single(existing.TextBoxes));
        Assert.Equal(initialNotes, NotesEntries(Save(deck, "after-failed-slide.pptx")));
        Invoke(runspace, """
            PowerPointSlide -Document $deck -Title 'Recovered' {
                PowerPointNotes 'Original notes'
                try {
                    PowerPointColumn -ErrorAction Stop {
                        PowerPointNotes 'Discarded column notes'
                        PowerPointParagraph 'Discarded one' -PassThru | Out-Null
                        PowerPointParagraph 'Discarded two'
                        PowerPointParagraph 'Discarded three'
                    }
                } catch { $script:columnFailed = $true }
                $script:accepted = PowerPointParagraph 'Accepted after failure' -PassThru
            }
            """);
        Assert.True(Get<bool>(runspace, "columnFailed"));
        AssertBacked(deck, Get<PowerPointParagraph>(runspace, "accepted"));
        using var saved = PowerPointPresentation.Load(Save(deck, "rollback.pptx"));
        Assert.Equal("Original notes", saved.Slides[1].GetSpeakerNotesText());
        string text = string.Join("\n", saved.Slides.SelectMany(s => s.TextBoxes).Select(b => b.Text));
        Assert.Contains("Existing content", text);
        Assert.Contains("Accepted after failure", text);
        Assert.DoesNotContain("Discarded", text);
        Assert.DoesNotContain("Rejected", text);
        Assert.Empty(saved.ValidateDocument());
    }

    [Fact]
    public void NestedRootsCannotBindOuterRunsAndRestoreTheOuterParagraphAndNativeSlide() {
        string inner = Path.Combine(_root, "inner.pptx");
        using var deck = PowerPointPresentation.Create();
        using var runspace = Open(deck);
        runspace.SessionStateProxy.SetVariable("inner", inner);
        Invoke(runspace, """
            PowerPointSlide -Document $deck -Title 'Outer' {
                $script:outerParagraph = PowerPointParagraph -PassThru {
                    PowerPointText 'Before '
                    New-OfficePowerPoint -Path $inner {
                        try { PowerPointText 'Must not leak' -ErrorAction Stop } catch { $script:isolated = $true }
                        PowerPointSlide -Title 'Inner' {
                            $script:innerParagraph = PowerPointParagraph 'Private' -PassThru
                            $script:innerNotes = PowerPointNotes 'Inner notes' -PassThru
                            $innerParagraph.Text = 'Private edited'
                        }
                    }
                    PowerPointText 'After'
                }
                $script:outerNotes = PowerPointNotes 'Outer notes' -PassThru
            }
            $outerParagraph.AddText(' saved') | Out-Null
            """);
        Assert.True(Get<bool>(runspace, "isolated"));
        AssertBacked(deck, Get<PowerPointParagraph>(runspace, "outerParagraph"));
        Assert.Same(deck.Slides[0], Get<PowerPointSlide>(runspace, "outerNotes"));
        using var saved = PowerPointPresentation.Load(Save(deck, "outer.pptx"));
        using var nested = PowerPointPresentation.Load(inner);
        string outerText = string.Join("\n", saved.Slides.SelectMany(s => s.TextBoxes).Select(b => b.Text));
        string innerText = string.Join("\n", nested.Slides.SelectMany(s => s.TextBoxes).Select(b => b.Text));
        Assert.Contains("Before After saved", outerText);
        Assert.DoesNotContain("Private", outerText);
        Assert.DoesNotContain("Must not leak", outerText + innerText);
        Assert.Contains("Private edited", innerText);
        Assert.DoesNotContain("Before", innerText);
        Assert.Equal("Outer notes", saved.Slides[0].GetSpeakerNotesText());
        Assert.Equal("Inner notes", nested.Slides[0].GetSpeakerNotesText());
        Assert.Empty(saved.ValidateDocument());
        Assert.Empty(nested.ValidateDocument());
    }

    private static string[] NotesEntries(string path) {
        using var archive = System.IO.Compression.ZipFile.OpenRead(path);
        return archive.Entries.Where(e => e.FullName.StartsWith("ppt/notes", StringComparison.Ordinal)).Select(e => {
            using var stream = e.Open();
            return e.FullName + ":" + Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(stream));
        }).OrderBy(s => s, StringComparer.Ordinal).ToArray();
    }

    private string Save(PowerPointPresentation deck, string name) {
        string path = Path.Combine(_root, name);
        deck.Save(path);
        return path;
    }
    private static object Node(PowerPointParagraph paragraph) => typeof(PowerPointParagraph)
        .GetProperty("Paragraph", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(paragraph)!;
    private static object Node(PowerPointTextRun run) => typeof(PowerPointTextRun)
        .GetProperty("Run", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(run)!;
    private static void AssertBacked(PowerPointPresentation deck, PowerPointParagraph paragraph) =>
        Assert.Contains(deck.Slides.SelectMany(s => s.TextBoxes).SelectMany(b => b.Paragraphs), p => ReferenceEquals(Node(p), Node(paragraph)));
    private static T Get<T>(Runspace runspace, string name) => Assert.IsType<T>(PSObject.AsPSObject(runspace.SessionStateProxy.GetVariable(name)).BaseObject);
    private static void Invoke(Runspace runspace, string script) {
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript(script);
        shell.Invoke();
        Assert.Equal(PSInvocationState.Completed, shell.InvocationStateInfo.State);
        Assert.Empty(shell.Streams.Error);
    }
    private static Runspace Open(PowerPointPresentation deck) {
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
        runspace.SessionStateProxy.SetVariable("deck", deck);
        return runspace;
    }
}
