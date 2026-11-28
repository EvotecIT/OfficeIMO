using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using Xunit;
using static OfficeIMO.Tests.PowerPointUnchangedSaveTests;

namespace OfficeIMO.Tests;

/// <summary>
/// Removing and moving slides in a presentation whose Open XML SDK package is not open yet edits the slide list without
/// opening it, and saves the package the slide model saves.
/// </summary>
public class PowerPointSlideOrderWithoutSdkTests {
    public static IEnumerable<object[]> CorpusPresentations() => OpenXmlSlideTextReaderTests.CorpusPresentations();

    private static readonly (string Name, int MinimumSlides, Action<PowerPointDocument, int> Edit)[] Edits = {
        ("remove the first slide", 1, (presentation, count) => presentation.RemoveSlide(0)),
        ("remove the last slide", 2, (presentation, count) => presentation.RemoveSlide(count - 1)),
        ("remove a middle slide", 3, (presentation, count) => presentation.RemoveSlide(count / 2)),
        ("move the first slide last", 2, (presentation, count) => presentation.MoveSlide(0, count - 1)),
        ("move the last slide first", 2, (presentation, count) => presentation.MoveSlide(count - 1, 0)),
        ("remove the first slide, then move the last one first", 3, (presentation, count) => {
            presentation.RemoveSlide(0);
            presentation.MoveSlide(count - 2, 0);
        }),
        ("add a slide", 0, (presentation, count) => presentation.AddSlide()),
        ("add a slide with the second layout", 0, (presentation, count) => presentation.AddSlide(0, 1)),
        ("remove the last slide, then add two", 1, (presentation, count) => {
            presentation.RemoveSlide(count - 1);
            presentation.AddSlide();
            presentation.AddSlide();
        }),
        ("duplicate the first slide", 1, (presentation, count) => presentation.DuplicateSlide(0)),
        ("duplicate the last slide to the front", 2, (presentation, count) => presentation.DuplicateSlide(count - 1, 0)),
        ("add a slide, then move it first", 1, (presentation, count) => {
            presentation.AddSlide();
            presentation.MoveSlide(count, 0);
        })
    };

    [Theory]
    [MemberData(nameof(CorpusPresentations))]
    public void EditsCorpusPresentationsAsTheModelDoes(string relativePath) {
        byte[] source = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, relativePath));
        AssertNoDifferences(Differences(source));
    }

    [Fact]
    public void RemovesAndMovesSlidesWithoutOpeningTheSdk() {
        byte[] source = CreateDeck();
        using var presentation = PowerPointDocument.Load(new MemoryStream(source));

        presentation.RemoveSlide(1);
        presentation.MoveSlide(2, 0);

        Assert.Equal(3, presentation.Slides.Count);
        string text = string.Concat(presentation.ExtractMarkdownChunks().Select(chunk => chunk.Text));
        byte[] saved = presentation.ToBytes();
        Assert.False(presentation.IsSdkLoaded);
        Assert.True(text.IndexOf("Fourth", StringComparison.Ordinal) < text.IndexOf("First", StringComparison.Ordinal));
        Assert.DoesNotContain("Second", text);
        Assert.DoesNotContain(ReadParts(saved).Keys, part => part.Contains("notesSlide", StringComparison.Ordinal));
        AssertNoDifferences(Differences(source));
    }

    [Fact]
    public void KeepsSectionsInLineWithTheSlides() {
        byte[] source;
        using (var created = PowerPointDocument.Load(new MemoryStream(CreateDeck()))) {
            created.AddSection("Opening", 0);
            created.AddSection("Middle", 1);
            created.AddSection("Closing", 3);
            source = created.ToBytes();
        }

        using var presentation = PowerPointDocument.Load(new MemoryStream(source));
        presentation.MoveSlide(0, 2);
        presentation.RemoveSlide(3);

        Assert.False(presentation.IsSdkLoaded);
        Assert.Equal(new[] { "Middle", "Opening" }, presentation.GetSections().Select(section => section.Name));
        AssertNoDifferences(Differences(source));
    }

    [Fact]
    public void TheSlidesAfterAnEditAreTheEditedOnes() {
        using var presentation = PowerPointDocument.Load(new MemoryStream(CreateDeck()));
        presentation.RemoveSlide(0);
        presentation.MoveSlide(0, 2);

        Assert.Equal(new[] { "Third slide", "Fourth slide", "Second slide" },
            presentation.Slides.Select(slide => slide.TextBoxes.Single().Text));
        Assert.False(presentation.IsSdkLoaded);
    }

    [Fact]
    public void ARemovedSlideACallerKeepsDoesNotReachASlideThatReusesItsPart() {
        // Slide parts 1, 3 and 4: the next new slide part after removing slide 4 is slide4.xml again.
        byte[] gapped;
        using (var source = PowerPointDocument.Load(new MemoryStream(CreateDeck()))) {
            source.RemoveSlide(1);
            gapped = source.ToBytes();
        }
        using var presentation = PowerPointDocument.Load(new MemoryStream(gapped));
        PowerPointSlide removed = presentation.Slides[2];
        presentation.RemoveSlide(2);
        PowerPointSlide copy = presentation.DuplicateSlide(0);
        Assert.False(presentation.IsSdkLoaded);
        Assert.Contains("ppt/slides/slide4.xml", ReadParts(presentation.ToBytes()).Keys);

        Record.Exception(() => removed.Hidden = true);

        Assert.False(copy.Hidden);
        Assert.Equal("First slide", copy.TextBoxes.Single().Text);
    }

    [Theory]
    [InlineData("custom show", false)]
    [InlineData("comment", true)]
    [InlineData("link", false)]
    public void SlideRemovalMatchesTheModelForRelatedContent(string variant, bool opensSdk) {
        byte[] source;
        using (var created = PowerPointDocument.Load(new MemoryStream(CreateDeck()))) {
            if (variant == "custom show") created.AddCustomShow("Short", new[] { created.Slides[0], created.Slides[2] });
            if (variant == "comment") created.AddClassicComment(created.Slides[0], new PowerPointCommentAuthor("Reviewer", "RV"), "Check this");
            if (variant == "link") created.Slides[1].TextBoxes.Single().SetHyperlink(created.Slides[0]);
            source = created.ToBytes();
        }

        using var presentation = PowerPointDocument.Load(new MemoryStream(source));
        presentation.RemoveSlide(0);

        Assert.Equal(opensSdk, presentation.IsSdkLoaded);
        AssertNoDifferences(Differences(source, Edits[0]));
    }

    [Fact]
    public void SlideRemovalUnlinksShapesRunsAndUnusedEdgesInTheOwnedPackage() {
        using var source = PowerPointDocument.Create();
        PowerPointSlide first = source.AddSlide();
        PowerPointSlide target = source.AddSlide();
        PowerPointSlide text = source.AddSlide();
        PowerPointSlide unused = source.AddSlide();
        PowerPointAutoShape linkedShape = first.AddRectangle(OfficeLength.Points(20), OfficeLength.Points(20),
            OfficeLength.Points(120), OfficeLength.Points(40));
        linkedShape.SetHyperlink(target);
        PowerPointAutoShape keptShape = first.AddRectangle(OfficeLength.Points(20), OfficeLength.Points(80),
            OfficeLength.Points(120), OfficeLength.Points(40));
        keptShape.SetHyperlink(text);
        PowerPointTextBox textBox = text.AddTextBox("Go to the removed slide");
        PowerPointTextRun run = textBox.Paragraphs.Single().Runs.Single();
        run.SetHyperlink(target);
        target.AddTextBox("Return").Paragraphs.Single().Runs.Single().SetHyperlink(first);
        PowerPointCustomShow show = source.AddCustomShow("Only target", new[] { target });

        A.HyperlinkOnClick combinedAction = ((P.Shape)linkedShape.Element).NonVisualShapeProperties!
            .NonVisualDrawingProperties!.GetFirstChild<A.HyperlinkOnClick>()!;
        combinedAction.Action = "ppaction://customshow?id=" + show.Id;
        string firstRelationship = combinedAction.Id!.Value!;
        string runRelationship = text.SlidePart.GetIdOfPart(target.SlidePart);
        P.TextBody textBody = ((P.Shape)textBox.Element).TextBody!;
        textBody.ListStyle ??= new A.ListStyle();
        textBody.ListStyle.Append(new A.Level1ParagraphProperties(new A.DefaultRunProperties(new A.HyperlinkOnClick {
            Id = runRelationship,
            Action = "ppaction://hlinksldjump"
        })));
        const string unusedRelationship = "rIdUnusedRemovedSlide";
        unused.SlidePart.AddPart(target.SlidePart, unusedRelationship);
        string firstPart = first.SlidePart.Uri.OriginalString;
        string targetPart = target.SlidePart.Uri.OriginalString;
        string textPart = text.SlidePart.Uri.OriginalString;
        string unusedPart = unused.SlidePart.Uri.OriginalString;

        using var input = OfficeOpcEditablePackage.Open(new MemoryStream(source.ToBytes(), writable: false));
        OfficeXmlTree textTree = input.GetXml(textPart);
        var textRoot = new OfficeXmlNode(textTree, textTree.Root);
        const int p = OfficeXmlNamespaces.PresentationML, a = OfficeXmlNamespaces.DrawingML, r = OfficeXmlNamespaces.Relationships;
        OfficeXmlNode extensions = textRoot.Child(p, "extLst");
        if (extensions.IsNone) extensions = textRoot.Append(textRoot.CreateElement(p, "extLst"));
        OfficeXmlNode payload = extensions.Append(extensions.CreateElement(p, "ext"));
        payload.SetAttribute("uri", "urn:fixture:inbound-link-decoy");
        OfficeXmlNode decoy = payload.Append(payload.CreateElement(a, "hlinkClick"));
        decoy.SetAttribute(r, "id", runRelationship);
        using var packaged = new MemoryStream();
        input.Save(packaged);
        byte[] authored = packaged.ToArray();

        using var presentation = PowerPointDocument.Load(new MemoryStream(authored, writable: false));
        PowerPointTextRun heldRun = presentation.Slides[2].TextBoxes.Single().Paragraphs.Single().Runs.Single();
        PowerPointSlide heldTarget = presentation.Slides[1];
        presentation.RemoveSlide(1);
        Assert.Empty(presentation.CustomShows);
        Assert.DoesNotContain(heldTarget, presentation.Slides);
        Assert.Null(heldRun.Hyperlink);
        byte[] output = presentation.ToBytes();
        Assert.False(presentation.IsSdkLoaded);

        using var saved = OfficeOpcEditablePackage.Open(new MemoryStream(output, writable: false));
        Assert.False(saved.PartExists(targetPart));
        Assert.DoesNotContain(firstRelationship, ReadRelationshipIds(saved, firstPart));
        Assert.DoesNotContain(runRelationship, ReadRelationshipIds(saved, textPart));
        Assert.DoesNotContain(unusedRelationship, ReadRelationshipIds(saved, unusedPart));
        OfficeXmlTree savedText = saved.GetXml(textPart);
        var savedRoot = new OfficeXmlNode(savedText, savedText.Root);
        Assert.Empty(savedRoot.ContentDescendants(a, "hlinkClick"));
        Assert.Equal(runRelationship, savedRoot.Child(p, "extLst").Child(p, "ext").Child(a, "hlinkClick").Attribute(r, "id"));

        OfficeXmlTree firstTree = saved.GetXml(firstPart);
        var firstRoot = new OfficeXmlNode(firstTree, firstTree.Root);
        Assert.Single(firstRoot.ContentDescendants(a, "hlinkClick"));
        AssertNoDifferences(Differences(authored, ("remove the target slide", 2, (deck, _) => deck.RemoveSlide(1))));

        static IEnumerable<string?> ReadRelationshipIds(OfficeOpcEditablePackage package, string part) {
            OfficeXmlTree tree = package.GetRelationships(part);
            return new OfficeXmlNode(tree, tree.Root).Children(OfficeXmlNamespaces.PackageRelationships, "Relationship").ToList()
                .Select(relationship => relationship.Attribute("Id"));
        }
    }

    [Fact]
    public void SlideRemovalPreservesSdkDeletionWhenAHyperlinkReferencesAnImagePart() {
        using var source = PowerPointDocument.Create();
        PowerPointSlide first = source.AddSlide();
        PowerPointSlide removed = source.AddSlide();
        first.AddTextBox("Go to the removed slide").SetHyperlink(removed);
        ImagePart image = removed.SlidePart.AddImagePart(ImagePartType.Png);
        using (var asset = new MemoryStream(Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=="))) {
            image.FeedData(asset);
        }
        string imageName = image.Uri.OriginalString;
        first.SlidePart.AddHyperlinkRelationship(new Uri(OfficeOpcPartNames.GetRelativeTarget(first.SlidePart.Uri.OriginalString,
            imageName), UriKind.Relative), false, "rIdImageReference");

        byte[] bytes = source.ToBytes();
        using var presentation = PowerPointDocument.Load(new MemoryStream(bytes, writable: false));
        Assert.False(presentation.IsSdkLoaded);
        presentation.RemoveSlide(1);

        Assert.True(presentation.IsSdkLoaded);
        using var saved = OfficeOpcPackage.Open(new MemoryStream(presentation.ToBytes(), writable: false));
        Assert.False(saved.PartExists(imageName));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SlideRemovalRetainsSdkValidationForMalformedRelationshipIdentifiers(bool duplicateOnUnlinkedSlide) {
        using var source = PowerPointDocument.Create();
        PowerPointSlide first = source.AddSlide();
        PowerPointSlide removed = source.AddSlide();
        PowerPointSlide unrelated = source.AddSlide();
        first.SlidePart.AddPart(removed.SlidePart, "rIdUnusedInbound");
        unrelated.SlidePart.AddHyperlinkRelationship(new Uri("https://example.test/one"), true, "rIdExternalOne");
        unrelated.SlidePart.AddHyperlinkRelationship(new Uri("https://example.test/two"), true, "rIdExternalTwo");
        using var edit = OfficeOpcEditablePackage.Open(new MemoryStream(source.ToBytes(), writable: false));
        OfficeXmlTree tree = edit.GetRelationships((duplicateOnUnlinkedSlide ? unrelated : first).SlidePart.Uri.OriginalString);
        string originalId = duplicateOnUnlinkedSlide ? "rIdExternalTwo" : "rIdUnusedInbound";
        const int ns = OfficeXmlNamespaces.PackageRelationships;
        var relationshipRoot = new OfficeXmlNode(tree, tree.Root);
        OfficeXmlNode edge = Assert.Single(relationshipRoot.Children(ns, "Relationship").ToList(),
            node => node.Attribute("Id") == originalId);
        edge.SetAttribute("Id", duplicateOnUnlinkedSlide ? "rIdExternalOne" : "bad id");
        using var corrupt = new MemoryStream();
        edit.Save(corrupt);
        byte[] bytes = corrupt.ToArray();

        using var sdk = PowerPointDocument.Load(new MemoryStream(bytes, writable: false));
        Exception expected = Assert.ThrowsAny<Exception>(() => _ = sdk.OpenXmlDocument);
        using var presentation = PowerPointDocument.Load(new MemoryStream(bytes, writable: false));
        Assert.False(presentation.IsSdkLoaded);
        Exception actual = Assert.ThrowsAny<Exception>(() => presentation.RemoveSlide(1));
        Assert.IsType(expected.GetType(), actual);
    }

    // Opt-in, as the other corpus checks: OFFICEIMO_PPTX_CORPUS names a folder of presentations.
    [Fact]
    public void ExternalCorpusPresentationsEditAsTheModelDoes() {
        string? folder = Environment.GetEnvironmentVariable("OFFICEIMO_PPTX_CORPUS");
        if (string.IsNullOrEmpty(folder) || !Directory.Exists(folder)) return;
        var failures = new List<string>();
        foreach (string path in Directory.EnumerateFiles(folder, "*.pptx", SearchOption.AllDirectories).OrderBy(file => file, StringComparer.Ordinal)) {
            byte[] source = File.ReadAllBytes(path);
            try {
                using (PowerPointDocument.Load(new MemoryStream(source))) { }
            } catch (Exception) {
                continue;
            }
            failures.AddRange(Differences(source).Select(difference => path + ": " + difference));
        }
        Assert.True(failures.Count == 0, string.Join(Environment.NewLine, failures.Take(40)));
    }

    private static List<string> Differences(byte[] source, params (string Name, int MinimumSlides, Action<PowerPointDocument, int> Edit)[] edits) {
        var differences = new List<string>();
        int count;
        using (var counted = PowerPointDocument.Load(new MemoryStream(source))) count = counted.Slides.Count;
        foreach ((string name, int minimumSlides, Action<PowerPointDocument, int> edit) in edits.Length == 0 ? Edits : edits) {
            if (count < minimumSlides) continue;
            byte[]? model = null, deferred = null;
            Exception? modelFailure = Record.Exception(() => {
                using var opened = PowerPointDocument.Load(new MemoryStream(source));
                _ = opened.OpenXmlDocument;
                edit(opened, count);
                model = opened.ToBytes();
            });
            Exception? deferredFailure = Record.Exception(() => {
                using var loaded = PowerPointDocument.Load(new MemoryStream(source));
                edit(loaded, count);
                deferred = loaded.ToBytes();
            });
            if (modelFailure != null || deferredFailure != null) {
                if (modelFailure?.GetType() != deferredFailure?.GetType()) differences.Add(name + ": model " + modelFailure?.GetType().Name + ", without the SDK " + deferredFailure?.GetType().Name);
                continue;
            }
            Exception? packageDifference = Record.Exception(() => AssertSamePackage(source, model!, deferred!, keptAsLoaded: false));
            if (packageDifference != null) differences.Add(name + ": " + packageDifference.Message);
        }
        return differences;
    }

    private static byte[] CreateDeck() {
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        presentation.AddSlide().AddTextBox("First slide");
        PowerPointSlide second = presentation.AddSlide();
        second.AddTextBox("Second slide");
        second.Notes.Text = "Speaker note";
        presentation.AddSlide().AddTextBox("Third slide");
        presentation.AddSlide().AddTextBox("Fourth slide");
        return presentation.ToBytes();
    }

    // The whole of each difference, which the collection display would cut short.
    private static void AssertNoDifferences(List<string> differences) =>
        Assert.True(differences.Count == 0, string.Join(Environment.NewLine, differences));
}
