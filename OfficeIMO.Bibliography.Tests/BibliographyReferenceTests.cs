using System.Collections.Generic;
using System.Text.Json;

namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReferenceTests {
    [Fact]
    public void CrossrefFillsMissingGroupsAndMapsTheContainingTitleWithoutChangingTheSource() {
        const string text = "@incollection{chapter, title={Chapter}, publisher={}, date={2020}, crossref={book}}\n" +
            "@book{book, title={Collected work}, publisher={Press}, location={London}, date={2019-04}, " +
            "editor={Doe, Jane}, isbn={9780000000000}, vendor={opaque}, options={hidden}}";
        BibliographyDocument source = Parse(text);
        BibliographyReferenceResult result = source.ResolveReferences();
        BibliographyItem chapter = result.Document.Items[0];

        Assert.True(result.IsComplete);
        Assert.Equal("Chapter", chapter.Title);
        Assert.Equal("Collected work", chapter.ContainerTitle);
        Assert.Equal(string.Empty, chapter.Publisher);
        Assert.Equal("London", chapter.PublisherPlace);
        Assert.Equal(2020, Assert.Single(chapter.Dates).Year);
        Assert.Null(chapter.Dates[0].Month);
        Assert.Equal("Doe", Assert.Single(chapter.Contributors).Name.Family);
        Assert.Equal("9780000000000", chapter.GetIdentifier("ISBN"));
        Assert.Equal("opaque", Assert.Single(chapter.NativeFields, field => field.Name == "vendor").Value);
        Assert.DoesNotContain(chapter.NativeFields, field => field.Name == "options");
        BibliographyFieldProvenance title = Assert.Single(result.Provenance, field => field.ItemKey == "chapter" && field.Field == "container-title");
        Assert.Equal("title", title.SourceField);
        Assert.Equal(new[] { "chapter", "book" }, title.ReferencePath);
        Assert.Equal(1, title.SourceItemIndex);
        Assert.Equal(text, source.ToString());
        Assert.False(source.IsModified);
        Assert.Null(source.Items[0].ContainerTitle);
        Assert.False(result.Document.HasOriginalSource);

        BibliographyItem reopened = Parse(result.Document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical }).Content).Items[0];
        Assert.Equal("Chapter", reopened.Title);
        Assert.Equal("Collected work", reopened.ContainerTitle);
    }

    [Fact]
    public void CascadingXdataUsesListedReplacementPrecedenceAndRetainsTheUltimateOrigin() {
        BibliographyDocument source = Parse("@book{child, title={Own}, publisher={Own press}, xdata={middle,last}, crossref={parent}}\n" +
            "@xdata{base, publisher={Base press}, location={Oxford}, author={Doe, Jane}}\n" +
            "@xdata{middle, xdata={base}, title={Middle title}}\n" +
            "@xdata{last, title={Last title}}\n" +
            "@book{parent, title={Parent}, publisher={Fallback}, volume={2}}", BibliographyFormat.BibLatex);
        BibliographyReferenceResult result = source.ResolveReferences();
        BibliographyItem child = result.Document.Items[0];

        Assert.True(result.IsComplete);
        Assert.Equal("Last title", child.Title);
        Assert.Equal("Base press", child.Publisher);
        Assert.Equal("Oxford", child.PublisherPlace);
        Assert.Equal("2", child.Volume);
        Assert.Equal(2, result.CitationItems.Count);
        Assert.Equal(5, result.Document.Items.Count);
        BibliographyFieldProvenance publisher = Assert.Single(result.Provenance, field => field.ItemKey == "child" && field.Field == "publisher");
        Assert.Equal("base", publisher.SourceItemKey);
        Assert.Equal(new[] { "child", "middle", "base" }, publisher.ReferencePath);
        Assert.Equal("xdata", publisher.Relation);
        Assert.Equal("last", Assert.Single(result.Provenance, field => field.ItemKey == "child" && field.Field == "title").SourceItemKey);
        child.Contributors[0].Name.Family = "Changed";
        Assert.Equal("Doe", source.Items[1].Contributors[0].Name.Family);
        Assert.Equal("Doe", result.Document.Items[1].Contributors[0].Name.Family);
        Assert.Equal("Doe", result.Document.Items[2].Contributors[0].Name.Family);
        Assert.Throws<NotSupportedException>(() => ((IList<string>)publisher.ReferencePath).Add("other"));

        BibliographyReferenceResult fallback = source.ResolveReferences(new BibliographyReferenceOptions { XDataOverridesExistingFields = false });
        Assert.Equal("Own", fallback.Document.Items[0].Title);
        Assert.Equal("Own press", fallback.Document.Items[0].Publisher);
        Assert.Equal("Oxford", fallback.Document.Items[0].PublisherPlace);
    }

    [Fact]
    public void KeyAndReferenceFailuresAreDiagnosedWithoutChoosingAnAmbiguousParent() {
        BibliographyDocument source = Parse("@book{child, xdata={ordinary,missing}, crossref={duplicate}}\n" +
            "@book{duplicate, publisher={First}}\n@book{duplicate, publisher={Second}}\n@book{ordinary, publisher={Other}}", BibliographyFormat.BibLatex);
        BibliographyReferenceResult result = source.ResolveReferences();

        Assert.True(result.HasErrors);
        Assert.False(result.IsComplete);
        Assert.Null(result.Document.Items[0].Publisher);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "BIBREF002" && diagnostic.ItemKey == "child");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "BIBREF003" && diagnostic.Field == "xdata");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "BIBREF007" && diagnostic.Field == "xdata");
        Assert.Throws<NotSupportedException>(() => ((IList<BibliographyDiagnostic>)result.Diagnostics).Clear());

        var authored = new BibliographyDocument(BibliographyFormat.BibTex);
        authored.Items.Add(new BibliographyItem());
        authored.Items.Add(new BibliographyItem { Key = "Case", Publisher = "Parent" });
        var child = new BibliographyItem { Key = "child" };
        child.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.BibTex, "crossref", "case"));
        authored.Items.Add(child);
        BibliographyReferenceResult caseSensitive = authored.ResolveReferences();
        Assert.Contains(caseSensitive.Diagnostics, diagnostic => diagnostic.Code == "BIBREF001");
        Assert.Contains(caseSensitive.Diagnostics, diagnostic => diagnostic.Code == "BIBREF003");
        Assert.Null(caseSensitive.Document.Items[2].Publisher);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CyclicRecordsAndTheirDependentsRetainOriginalValuesRegardlessOfRecordOrder(bool reverse) {
        BibliographyDocument source = Parse("@book{a, title={A}, crossref={b}}\n@book{b, publisher={B}, crossref={a}}\n" +
            "@book{dependent, crossref={a}}\n@book{independent, crossref={root}}\n@book{root, publisher={Good}}");
        if (reverse) ReverseItems(source);
        BibliographyReferenceResult result = source.ResolveReferences();

        Assert.Equal(3, result.Diagnostics.Count(diagnostic => diagnostic.Code == "BIBREF004"));
        Assert.Null(result.Document.Items.Single(item => item.Key == "a").Publisher);
        Assert.Null(result.Document.Items.Single(item => item.Key == "b").Title);
        Assert.Null(result.Document.Items.Single(item => item.Key == "dependent").Title);
        Assert.Equal("Good", result.Document.Items.Single(item => item.Key == "independent").Publisher);
        Assert.DoesNotContain(result.Provenance, field => field.ItemKey == "a" || field.ItemKey == "b" || field.ItemKey == "dependent");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ChainDepthIsAPropertyOfTheGraphAndCannotBeBypassedByMemoizedParents(bool reverse) {
        BibliographyDocument source = Parse("@book{a, crossref={b}}\n@book{b, crossref={c}}\n@book{c, crossref={d}}\n" +
            "@book{d, crossref={root}}\n@book{root, publisher={Press}}");
        if (reverse) ReverseItems(source);
        BibliographyReferenceResult result = source.ResolveReferences(new BibliographyReferenceOptions { MaximumDepth = 2 });

        Assert.Equal("Press", result.Document.Items.Single(item => item.Key == "c").Publisher);
        Assert.Null(result.Document.Items.Single(item => item.Key == "b").Publisher);
        Assert.Null(result.Document.Items.Single(item => item.Key == "a").Publisher);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "BIBREF005" && diagnostic.ItemKey == "b");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "BIBREF008" && diagnostic.ItemKey == "a");
    }

    [Fact]
    public void SnapshotCopyKeepsNumericIdentityAndDoesNotReviveEditedNativeRawValues() {
        BibliographyDocument source = BibliographyDocument.Parse("{\"id\":14.0,\"type\":\"book\",\"title\":\"Work\",\"vendor\":3," +
            "\"author\":[{\"family\":\"Doe\",\"vendor\":true}]}", BibliographyFormat.CslJson).Document;
        source.Items[0].NativeFields[0].Value = "7";
        BibliographyReferenceResult result = source.ResolveReferences();
        result.Document.Items[0].Contributors[0].Name.NativeFields[0].Value = "false";
        string content = result.Document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical }).Content;
        using JsonDocument json = JsonDocument.Parse(content);

        Assert.Equal(JsonValueKind.Object, json.RootElement.ValueKind);
        Assert.Equal(JsonValueKind.Number, json.RootElement.GetProperty("id").ValueKind);
        Assert.Equal(14m, json.RootElement.GetProperty("id").GetDecimal());
        Assert.Equal("7", json.RootElement.GetProperty("vendor").ToString());
        Assert.Equal("false", json.RootElement.GetProperty("author")[0].GetProperty("vendor").ToString().ToLowerInvariant());
        Assert.Equal("true", source.Items[0].Contributors[0].Name.NativeFields[0].Value);
    }

    [Fact]
    public void RepeatedReferenceFieldsAndEmptyReferenceEntriesHaveExplicitDiagnostics() {
        BibliographyReferenceResult result = Parse("@book{child, crossref={first}, crossref={second}, xdata={,}}\n" +
            "@book{first, publisher={First}}\n@book{second, publisher={Second}}", BibliographyFormat.BibLatex).ResolveReferences();
        Assert.Null(result.Document.Items[0].Publisher);
        Assert.Equal(3, result.Diagnostics.Count(diagnostic => diagnostic.Code == "BIBREF006"));
        Assert.True(result.HasErrors);
    }

    [Fact]
    public void ResolutionCanBeDisabledAndCancellationAndExpansionLimitsPreserveTheInput() {
        BibliographyDocument source = Parse("@book{child, crossref={missing}, xdata={missing}}", BibliographyFormat.BibLatex);
        BibliographyReferenceResult disabled = source.ResolveReferences(new BibliographyReferenceOptions { ResolveCrossref = false, ResolveXData = false, MaximumDepth = 0 });
        Assert.True(disabled.IsComplete);
        Assert.Empty(disabled.Provenance);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => source.ResolveReferences(cancellationToken: cancellation.Token));
        Assert.Throws<ArgumentOutOfRangeException>(() => source.ResolveReferences(new BibliographyReferenceOptions { MaximumDepth = -1 }));
        Assert.Throws<InvalidOperationException>(() => source.ResolveReferences(new BibliographyReferenceOptions { MaximumReferences = 1 }));

        var expanded = new BibliographyDocument(BibliographyFormat.BibLatex);
        expanded.Items.Add(new BibliographyItem { Key = "base", NativeType = "xdata", Publisher = new string('p', 100) });
        for (int index = 0; index < 15; index++) {
            var item = new BibliographyItem { Key = "child" + index, Type = BibliographyItemType.Book };
            item.NativeFields.Add(new BibliographyNativeField(BibliographyFormat.BibLatex, "xdata", "base"));
            expanded.Items.Add(item);
        }
        Assert.Throws<InvalidOperationException>(() => expanded.ResolveReferences(new BibliographyReferenceOptions { MaximumExpandedCharacters = 1_000 }));
        Assert.All(expanded.Items.Skip(1), item => Assert.Null(item.Publisher));
        Assert.Throws<InvalidOperationException>(() => expanded.ResolveReferences(new BibliographyReferenceOptions { MaximumItems = 5 }));
        Assert.Throws<InvalidOperationException>(() => expanded.ResolveReferences(new BibliographyReferenceOptions { MaximumValues = 5 }));
    }

    private static BibliographyDocument Parse(string text, BibliographyFormat format = BibliographyFormat.BibTex) => BibliographyDocument.Parse(text, format).Document;

    private static void ReverseItems(BibliographyDocument document) {
        BibliographyItem[] reversed = document.Items.Reverse().ToArray();
        document.Items.Clear();
        foreach (BibliographyItem item in reversed) document.Items.Add(item);
    }
}
