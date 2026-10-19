using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfTaggingStandardsTests {
    [Theory]
    [InlineData(PdfObjectSerializationMode.Buffered)]
    [InlineData(PdfObjectSerializationMode.ForwardOnly)]
    public void Pdf20DeclaresOneDocumentAndAnIndirectNamespaceOnEveryElement(PdfObjectSerializationMode mode) {
        var bytes = PdfDocument.Create(new PdfOptions {
            FileVersion = PdfFileVersion.Pdf20,
            TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers,
            Language = "en-US",
            ObjectSerializationMode = mode
        }.UseFontFamily("Carlito", PdfComplianceTestFonts.FindBundledTrueTypeFont()!)
            .UsePdfUa(PdfComplianceProfile.PdfUa2)).Meta(title: "Tagged standards")
            .Paragraph(paragraph => paragraph.Text("Accessible document")).ToBytes();
        var raw = PdfReadDocument.Open(bytes).RawStructure();
        var root = ObjectOfType(raw, "StructTreeRoot");
        var document = Assert.Single(Elements(raw), item => item.Value.Entries["S"].Text == "Document");
        Assert.Equal(document.ObjectNumber, Assert.Single(root.Value.Entries["K"].Items).ReferenceObjectNumber);
        int nsId = Assert.Single(root.Value.Entries["Namespaces"].Items).ReferenceObjectNumber!.Value;
        var ns = raw.GetObject(nsId)!.Value;
        Assert.Equal("Namespace", ns.Entries["Type"].Text);
        Assert.Equal("http://iso.org/pdf2/ssn", ns.Entries["NS"].Text);
        Assert.All(Elements(raw), item => Assert.Equal(nsId, item.Value.Entries["NS"].ReferenceObjectNumber));
        Assert.Equal("en-US", document.Value.Entries["Lang"].Text);
        var tagged = Assert.IsType<PdfTaggedContentInfo>(PdfInspector.Inspect(bytes).TaggedContent);
        var readback = PdfComplianceAnalyzer.AssessReadback(PdfComplianceProfile.PdfUa2, bytes);
        Assert.Equal(PdfComplianceRequirementStatus.Satisfied, readback.Requirements.Single(requirement =>
            requirement.Id == "readback-document-structure-element").Status);
        Assert.Single(tagged.RootElementObjectNumbers);
        Assert.Single(tagged.StructureElements, element => element.StructureType == "Document");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TableCellsResolveUniqueHeaderIdsThroughTheIdTree(bool containerCells) {
        var document = PdfDocument.Create().TaggedPdfCatalogMarkers();
        if (containerCells) {
            document.Canvas(canvas => canvas.Structure(PdfCanvasStructureRole.Table, table => table
                .Structure(PdfCanvasStructureRole.TableRow, row => row
                    .Structure(PdfCanvasStructureRole.TableHeaderCell, cell => cell.Text("Group", 20, 20, 100, 20),
                        new PdfCanvasStructureOptions { ColumnSpan = 2, HeaderScope = PdfCanvasTableHeaderScope.Column }))
                .Structure(PdfCanvasStructureRole.TableBody, body => body
                    .Structure(PdfCanvasStructureRole.TableRow, row => row
                        .Structure(PdfCanvasStructureRole.TableHeaderCell, cell => cell.Text("Row", 20, 60, 100, 20),
                            new PdfCanvasStructureOptions { HeaderScope = PdfCanvasTableHeaderScope.Row, RowSpan = 2 })
                        .Structure(PdfCanvasStructureRole.TableCell, cell => cell.Text("Value", 140, 60, 100, 20)))
                    .Structure(PdfCanvasStructureRole.TableRow, row => row
                        .Structure(PdfCanvasStructureRole.TableCell, cell => cell.Text("Next", 140, 100, 100, 20))))));
        } else {
            var style = TableStyles.Minimal();
            style.HeaderRowCount = 1;
            document.Table(new[] { new[] { "First", "Second" }, new[] { "One", "Two" } }, style: style);
        }
        var raw = PdfReadDocument.Open(document.ToBytes()).RawStructure();
        var headers = Elements(raw).Where(item => item.Value.Entries["S"].Text == "TH").ToArray();
        var ids = headers.Select(header => header.Value.Entries["ID"].Text).ToArray();
        Assert.Equal(2, headers.Length);
        Assert.Equal(2, ids.Distinct().Count());
        var root = ObjectOfType(raw, "StructTreeRoot");
        var names = raw.GetObject(root.Value.Entries["IDTree"].ReferenceObjectNumber!.Value)!.Value.Entries["Names"].Items;
        Assert.Equal(4, names.Count);
        for (int i = 0; i < names.Count; i += 2)
            Assert.Equal(names[i].Text, raw.GetObject(names[i + 1].ReferenceObjectNumber!.Value)!.Value.Entries["ID"].Text);
        var cells = Elements(raw).Where(item => item.Value.Entries["S"].Text == "TD").ToArray();
        Assert.NotEmpty(cells);
        foreach (var cell in cells) {
            var associations = cell.Value.Entries["A"].Entries["Headers"].Items.Select(item => item.Text).ToArray();
            if (containerCells) Assert.Equal(ids, associations);
            else Assert.Equal(ids[Array.IndexOf(cells, cell)], Assert.Single(associations));
            Assert.All(associations, id => Assert.Contains(id, ids));
        }
    }

    [Theory]
    [InlineData(130)]
    [InlineData(4100)]
    public void ParentNumberTreeHasIndirectBoundedNodesAndExactLimits(int count) {
        var entries = Enumerable.Range(0, count).Select(index => index % 2 == 0
            ? PdfStructTreeRootDictionaryBuilder.ParentTreeEntry.ForMarkedContentPage(index * 2, new int?[] { null, index + 10000 })
            : PdfStructTreeRootDictionaryBuilder.ParentTreeEntry.ForObjectReference(index * 2, index + 10000)).Reverse().ToArray();
        var nodes = new List<string>();
        string root = PdfStructTreeRootDictionaryBuilder.BuildParentTree(entries, value => { nodes.Add(value); return nodes.Count; });
        Assert.Contains("/Kids [", root);
        Assert.DoesNotContain("/Limits", root);
        Assert.All(nodes, node => Assert.Contains("/Limits [", node));
        // Follow every indirect child in order; limits must describe exactly its descendants.
        var seen = new List<int>();
        (int First, int Last) Visit(string node, bool isRoot) {
            var matches = System.Text.RegularExpressions.Regex.Matches(node, @"(?<!\d)(\d+) (?:\[null (\d+) 0 R\]|(\d+) 0 R)");
            int first, last;
            if (node.Contains("/Nums")) {
                Assert.InRange(matches.Count, 1, 64);
                foreach (System.Text.RegularExpressions.Match match in matches) {
                    int key = int.Parse(match.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
                    int objectId = int.Parse(match.Groups[2].Success ? match.Groups[2].Value : match.Groups[3].Value, System.Globalization.CultureInfo.InvariantCulture);
                    Assert.Equal(key / 2 + 10000, objectId);
                    seen.Add(key);
                }
                first = seen[seen.Count - matches.Count]; last = seen[seen.Count - 1];
            } else {
                var kids = System.Text.RegularExpressions.Regex.Matches(node.Split(new[] { "]" }, StringSplitOptions.None)[0], @"(\d+) 0 R");
                Assert.InRange(kids.Count, 1, 64);
                first = -1; last = -1;
                foreach (System.Text.RegularExpressions.Match kid in kids) {
                    var range = Visit(nodes[int.Parse(kid.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture) - 1], false);
                    if (first < 0) first = range.First;
                    last = range.Last;
                }
            }
            if (!isRoot) Assert.Contains("/Limits [" + first + " " + last + "]", node);
            return (first, last);
        }
        Visit(root, true);
        Assert.Equal(Enumerable.Range(0, count).Select(index => index * 2), seen);
    }

    [Fact]
    public void ParentTreePreservesEmptyArraysAndSparseMcidPositions() {
        var tree = PdfStructTreeRootDictionaryBuilder.BuildParentTree(new[] {
            PdfStructTreeRootDictionaryBuilder.ParentTreeEntry.ForMarkedContentPage(0, Array.Empty<int>()),
            PdfStructTreeRootDictionaryBuilder.ParentTreeEntry.ForMarkedContentPage(2, new int?[] { null, 9 }),
            PdfStructTreeRootDictionaryBuilder.ParentTreeEntry.ForObjectReference(4, 10)
        });
        Assert.Equal("<< /Nums [0 [] 2 [null 9 0 R] 4 10 0 R] >>\n", tree);
    }

    [Fact]
    public void LayoutLanguageIsAttachedToTheSpanWithoutChangingDocumentLanguage() {
        var document = PdfDocument.Create(new PdfOptions { Language = "en-US" }).TaggedPdfCatalogMarkers();
        document.Compose(compose => compose.Page(page => {
            page.Options.Language = "fr-FR";
            page.Content(content => content.Canvas(canvas => canvas.Text(new[] { PdfTextRun.Normal("Bonjour") }, PdfCanvasTextStructureRole.Span, 20, 20, 100, 20)));
        }));
        var raw = PdfReadDocument.Open(document.ToBytes()).RawStructure();
        var root = Assert.Single(Elements(raw), item => item.Value.Entries["S"].Text == "Document");
        Assert.Equal("en-US", root.Value.Entries["Lang"].Text);
        Assert.Contains(Elements(raw), item => item.Value.Entries["S"].Text == "Span" &&
            item.Value.Entries.TryGetValue("Lang", out var language) && language.Text == "fr-FR");
    }

    [Fact]
    public void PublicSpanLanguageIsSnapshottedInheritedAndCanReturnToDocumentLanguage() {
        var french = new PdfCanvasStructureOptions { Language = "fr-FR" };
        var document = PdfDocument.Create(new PdfOptions { Language = "en-US" }).TaggedPdfCatalogMarkers()
            .Canvas(canvas => canvas.Structure(PdfCanvasStructureRole.Paragraph, paragraph => paragraph
                .Structure(PdfCanvasStructureRole.Span, span => span
                    .Text(new[] { PdfTextRun.Normal("Bonjour") }, PdfCanvasTextStructureRole.Span, 20, 20, 100, 20)
                    .Structure(PdfCanvasStructureRole.Span, inherited => inherited
                        .Text(new[] { PdfTextRun.Normal("Encore") }, PdfCanvasTextStructureRole.Span, 120, 20, 100, 20))
                    .Structure(PdfCanvasStructureRole.Span, english => english
                        .Text(new[] { PdfTextRun.Normal("Hello") }, PdfCanvasTextStructureRole.Span, 220, 20, 100, 20),
                        new PdfCanvasStructureOptions { Language = "en-US" }), french)
                .Structure(PdfCanvasStructureRole.Span, sibling => sibling
                    .Text(new[] { PdfTextRun.Normal("After") }, PdfCanvasTextStructureRole.Span, 320, 20, 100, 20))));
        french.Language = "de-DE";
        var raw = PdfReadDocument.Open(document.ToBytes()).RawStructure();
        var elements = Elements(raw).ToDictionary(element => element.ObjectNumber);
        var root = Assert.Single(elements.Values, item => item.Value.Entries["S"].Text == "Document");
        var frenchSpan = Assert.Single(elements.Values, item => item.Value.Entries["S"].Text == "Span" &&
            item.Value.Entries.TryGetValue("Lang", out var language) && language.Text == "fr-FR");
        var englishSpan = Assert.Single(elements.Values, item => item.Value.Entries["S"].Text == "Span" &&
            item.Value.Entries.TryGetValue("Lang", out var language) && language.Text == "en-US");
        Assert.Equal(frenchSpan.ObjectNumber, englishSpan.Value.Entries["P"].ReferenceObjectNumber);
        string? EffectiveLanguage(PdfRawObjectView element) {
            while (true) {
                if (element.Value.Entries.TryGetValue("Lang", out var language)) return language.Text;
                if (!element.Value.Entries.TryGetValue("P", out var parent) || !parent.ReferenceObjectNumber.HasValue ||
                    !elements.TryGetValue(parent.ReferenceObjectNumber.Value, out var ancestor)) return null;
                element = ancestor;
            }
        }
        var textSpans = elements.Values.Where(element => element.Value.Entries["S"].Text == "Span" &&
            element.Value.Entries["K"].Kind == PdfRawValueKind.Dictionary).ToArray();
        Assert.Equal(4, textSpans.Length);
        Assert.Equal(new[] { "fr-FR", "fr-FR", "en-US", "en-US" }, textSpans.Select(EffectiveLanguage));
        Assert.Equal("en-US", root.Value.Entries["Lang"].Text);
        Assert.DoesNotContain(elements.Values, item => item.Value.Entries.TryGetValue("Lang", out var language) && language.Text == "de-DE");
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("fr\nFR")]
    [InlineData("fr\u0000FR")]
    [InlineData("en US")]
    [InlineData("en_US")]
    [InlineData("English")]
    [InlineData("en--US")]
    public void StructureLanguageRejectsInvalidDeclaredTags(string language) {
        Assert.Throws<ArgumentException>(() => new PdfCanvasStructureOptions { Language = language });
    }

    [Theory]
    [InlineData(null)]
    [InlineData("en")]
    [InlineData("fr-FR")]
    [InlineData("zh-Hant-TW")]
    public void StructureLanguageAcceptsProfileSyntaxAndInheritance(string? language) {
        var options = new PdfCanvasStructureOptions { Language = language };
        Assert.Equal(language, options.Language);
    }

    private static PdfRawObjectView ObjectOfType(PdfRawDocumentView raw, string type) =>
        Assert.Single(raw.Objects, item => item.Value.Entries.TryGetValue("Type", out var value) && value.Text == type);

    private static IEnumerable<PdfRawObjectView> Elements(PdfRawDocumentView raw) =>
        raw.Objects.Where(item => item.Value.Entries.TryGetValue("Type", out var type) && type.Text == "StructElem");
}
