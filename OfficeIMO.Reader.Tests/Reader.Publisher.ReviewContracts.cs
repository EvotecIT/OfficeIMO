using OfficeIMO.Core.Internal;
using OfficeIMO.Markdown;
using OfficeIMO.Pdf;
using OfficeIMO.Publisher;
using OfficeIMO.Reader.Publisher;
using System.Text;
using System.Text.RegularExpressions;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderPublisherReviewTests {
    [Theory]
    [InlineData("Simple.pub")]
    [InlineData("Sample.pub")]
    public void NativeStoryParagraphOrderSurvivesTransportAndPdfProjection(string fixture) {
        byte[] bytes = Fixture(fixture);
        OfficeDocumentReadResult source = new OfficeDocumentReaderBuilder().AddPublisherHandler().Build().ReadDocument(bytes, fixture);
        source = OfficeDocumentReadResultJson.Deserialize(OfficeDocumentReadResultJson.Serialize(source));
        string text = Normalize(PdfReadDocument.Open(source.ToPdfDocumentResult(new PdfProjectionOptions {
            PagePolicy = PdfProjectionPagePolicy.ContinuousFlow, IncludeMetadata = false
        }).ToBytes()).ExtractText());
        int position = 0;
        foreach (OfficeDocumentBlock block in source.Blocks) {
            string expected = Normalize(block.Text ?? string.Empty);
            if (expected.Length == 0) continue;
            int found = text.IndexOf(expected, position, StringComparison.Ordinal);
            Assert.True(found >= position, $"Source paragraph {block.Id} is missing or reordered after position {position}: {expected}");
            position = found + expected.Length;
        }
    }

    [Theory]
    [InlineData("    *value*")]
    [InlineData("\t*value*")]
    [InlineData("  \t[value](url)")]
    public void LiteralIndentationAndPunctuationRemainParagraphText(string text) {
        string markdown = ReaderMarkdownEscaping.EscapeLiteral(text);
        var paragraph = Assert.IsType<OfficeIMO.Markdown.ParagraphBlock>(Assert.Single(MarkdownReader.Parse(markdown).Blocks));
        var literal = new StringBuilder();
        ((IPlainTextMarkdownInline)paragraph.Inlines).AppendPlainText(literal);
        Assert.Equal(text, literal.ToString());
    }

    [Fact]
    public void IndentedNativeBulletBodyRemainsLiteralListContent() {
        const string text = "    *value*";
        var list = Assert.IsType<UnorderedListBlock>(Assert.Single(MarkdownReader.Parse("- " + ReaderMarkdownEscaping.EscapeLiteral(text)).Blocks));
        ListItem item = Assert.Single(list.Items);
        var literal = new StringBuilder();
        ((IPlainTextMarkdownInline)item.Content).AppendPlainText(literal);
        Assert.Equal(text, literal.ToString());
        Assert.Empty(item.NestedBlocks);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ContinuousPdfHonorsExplicitOrderAcrossPhysicalPagesAndTables(bool transport) {
        OfficeDocumentBlock Block(string text, int page, long order, int localIndex) => new() {
            Id = text, Text = text, Location = new() { Page = page, LogicalOrder = order, SourceBlockIndex = localIndex }
        };
        var first = Block("OrderAlpha", 2, 0, 0);
        var second = Block("OrderBravo", 2, 1, 1);
        var last = Block("OrderDelta", 1, 3, 0);
        var table = new ReaderTable { Columns = new[] { "OrderCharlie" }, Rows = new[] { new[] { "cell" } },
            Location = new() { Page = 2, LogicalOrder = 2, SourceBlockIndex = 0 } };
        var source = new OfficeDocumentReadResult { Kind = ReaderInputKind.Xps, Blocks = new[] { last, second, first }, Tables = new[] { table },
            Pages = new[] {
                new OfficeDocumentPage { Number = 1, Name = "PhysicalFirst", Blocks = new[] { last } },
                new OfficeDocumentPage { Number = 2, Name = "PhysicalSecond", Blocks = new[] { second, first }, Tables = new[] { table } }
            } };
        if (transport) source = OfficeDocumentReadResultJson.Deserialize(OfficeDocumentReadResultJson.Serialize(source));
        PdfDocumentConversionResult conversion = source.ToPdfDocumentResult(new PdfProjectionOptions {
            PagePolicy = PdfProjectionPagePolicy.ContinuousFlow, IncludeMetadata = false
        });
        string text = PdfReadDocument.Open(conversion.ToBytes()).ExtractText();
        Assert.Contains(conversion.Warnings, warning => warning.Code == "MODEL_SOURCE_PAGE_LABELS_OMITTED" && warning.LossKind == OfficeConversionLossKind.Omission);
        Assert.DoesNotContain("PhysicalFirst", text);
        Assert.True(text.IndexOf("OrderAlpha", StringComparison.Ordinal) < text.IndexOf("OrderBravo", StringComparison.Ordinal));
        Assert.True(text.IndexOf("OrderBravo", StringComparison.Ordinal) < text.IndexOf("OrderCharlie", StringComparison.Ordinal));
        Assert.True(text.IndexOf("OrderCharlie", StringComparison.Ordinal) < text.IndexOf("OrderDelta", StringComparison.Ordinal));
        Assert.Single(Regex.Matches(text, "OrderAlpha"));
        Assert.Single(Regex.Matches(text, "OrderCharlie"));
    }

    [Theory]
    [InlineData(1, 2, false)]
    [InlineData(1, 2, true)]
    [InlineData(2, 1, false)]
    [InlineData(2, 1, true)]
    [InlineData(2, 2, false)]
    [InlineData(2, 2, true)]
    public void PdfMergesAggregateAndPageTablesWithoutLosingEqualOccurrences(int aggregateCount, int pageCount, bool transport) {
        var table = new ReaderTable { Columns = new[] { "OccurrenceTable" }, Rows = new[] { new[] { "cell" } },
            Location = new() { Page = 1, LogicalOrder = 5 } };
        var source = new OfficeDocumentReadResult { Tables = Enumerable.Repeat(table, aggregateCount).ToArray(),
            Pages = new[] { new OfficeDocumentPage { Number = 1, Tables = Enumerable.Repeat(table, pageCount).ToArray() } } };
        if (transport) source = OfficeDocumentReadResultJson.Deserialize(OfficeDocumentReadResultJson.Serialize(source));
        string text = PdfReadDocument.Open(source.ToPdfDocumentResult(new PdfProjectionOptions {
            PagePolicy = PdfProjectionPagePolicy.ContinuousFlow, IncludeMetadata = false
        }).ToBytes()).ExtractText();
        Assert.Equal(Math.Max(aggregateCount, pageCount), Regex.Matches(text, "OccurrenceTable").Count);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void UniquePageTableWithoutAuthoredLocationsIsNotRepeatedByItsAggregate(bool transport) {
        var table = new ReaderTable { Columns = new[] { "UniquePageTable" }, Rows = new[] { new[] { "cell" } } };
        var source = new OfficeDocumentReadResult { Tables = new[] { table },
            Pages = new[] { new OfficeDocumentPage { Number = 1, Tables = new[] { table } } } };
        if (transport) source = OfficeDocumentReadResultJson.Deserialize(OfficeDocumentReadResultJson.Serialize(source));
        string text = PdfReadDocument.Open(source.ToPdfDocumentResult(new PdfProjectionOptions { IncludeMetadata = false }).ToBytes()).ExtractText();
        Assert.Single(Regex.Matches(text, "UniquePageTable"));
        Assert.Null(table.Location);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AmbiguousUnlocatedAggregateTableIsPreservedWithUnassessedEvidence(bool transport) {
        ReaderTable Table() => new() { Columns = new[] { "AmbiguousTable" }, Rows = new[] { new[] { "cell" } } };
        var source = new OfficeDocumentReadResult { Tables = new[] { Table() }, Pages = new[] {
            new OfficeDocumentPage { Number = 1, Tables = new[] { Table() } },
            new OfficeDocumentPage { Number = 2, Tables = new[] { Table() } }
        } };
        if (transport) source = OfficeDocumentReadResultJson.Deserialize(OfficeDocumentReadResultJson.Serialize(source));
        PdfDocumentConversionResult conversion = source.ToPdfDocumentResult(new PdfProjectionOptions { IncludeMetadata = false });
        Assert.Equal(3, Regex.Matches(PdfReadDocument.Open(conversion.ToBytes()).ExtractText(), "AmbiguousTable").Count);
        Assert.Contains(conversion.Warnings, warning => warning.Code == "MODEL_TABLE_CORRELATION_UNASSESSED"
            && warning.LossKind == OfficeConversionLossKind.Unassessed);
        Assert.Throws<InvalidOperationException>(() => conversion.RequireNoLoss());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void EqualTableContentWithConflictingSourcePositionsRemainsDistinct(bool transport) {
        ReaderTable Table(int page) => new() { Columns = new[] { "DistinctTable" }, Rows = new[] { new[] { "cell" } },
            Location = new() { Page = page, LogicalOrder = page } };
        var source = new OfficeDocumentReadResult { Tables = new[] { Table(2) },
            Pages = new[] { new OfficeDocumentPage { Number = 1, Tables = new[] { Table(1) } } } };
        if (transport) source = OfficeDocumentReadResultJson.Deserialize(OfficeDocumentReadResultJson.Serialize(source));
        PdfDocumentConversionResult conversion = source.ToPdfDocumentResult(new PdfProjectionOptions { IncludeMetadata = false });
        Assert.Equal(2, Regex.Matches(PdfReadDocument.Open(conversion.ToBytes()).ExtractText(), "DistinctTable").Count);
        Assert.DoesNotContain(conversion.Warnings, warning => warning.Code == "MODEL_TABLE_CORRELATION_UNASSESSED");
    }

    [Theory]
    [InlineData(0, false)]
    [InlineData(0, true)]
    [InlineData(1, false)]
    [InlineData(1, true)]
    [InlineData(2, false)]
    [InlineData(2, true)]
    [InlineData(3, false)]
    [InlineData(3, true)]
    public void TableOccurrenceMatchingRespectsKnownTextSpans(int coordinate, bool transport) {
        ReaderTable Table(int value) => new() {
            Columns = new[] { "SpanTable" }, Rows = new[] { new[] { "cell" } },
            Location = new() {
                Page = 1, StartLine = 0,
                EndLine = coordinate is 0 or 3 ? value : null,
                NormalizedStartLine = coordinate is 1 or 3 ? value : null,
                NormalizedEndLine = coordinate is 2 or 3 ? value + 2 : null
            }
        };
        bool equalSpan = coordinate == 3;
        var source = new OfficeDocumentReadResult { Tables = new[] { Table(equalSpan ? 10 : 20) },
            Pages = new[] { new OfficeDocumentPage { Number = 1, Tables = new[] { Table(10) } } } };
        if (transport) source = OfficeDocumentReadResultJson.Deserialize(OfficeDocumentReadResultJson.Serialize(source));
        PdfDocumentConversionResult conversion = source.ToPdfDocumentResult(new PdfProjectionOptions { IncludeMetadata = false });
        string text = PdfReadDocument.Open(conversion.ToBytes()).ExtractText();
        Assert.Equal(equalSpan ? 1 : 2, Regex.Matches(text, "SpanTable").Count);
        Assert.DoesNotContain(conversion.Warnings, warning => warning.Code == "MODEL_TABLE_CORRELATION_UNASSESSED");
    }

    [Fact]
    public void AuthoredNeutralLogicalPositionsRetainRepeatedParagraphsAroundTables() {
        var source = new OfficeDocumentModel { Blocks = new[] {
            new OfficeDocumentModelBlock { Text = "RepeatedParagraph", Location = new() { SourceBlockIndex = 0, LogicalOrder = 0 } },
            new OfficeDocumentModelBlock { Text = "RepeatedParagraph", Location = new() { SourceBlockIndex = 0, LogicalOrder = 2 } }
        }, Tables = new[] { new OfficeDocumentModelTable {
            Columns = new[] { "MiddleTable" }, Rows = new[] { new[] { "cell" } }, Location = new() { LogicalOrder = 1 }
        } } };
        string text = PdfReadDocument.Open(source.ToPdfDocumentResult(new PdfProjectionOptions { IncludeMetadata = false }).ToBytes()).ExtractText();
        Assert.Equal(2, Regex.Matches(text, "RepeatedParagraph").Count);
        Assert.True(text.IndexOf("RepeatedParagraph", StringComparison.Ordinal) < text.IndexOf("MiddleTable", StringComparison.Ordinal));
        Assert.True(text.IndexOf("MiddleTable", StringComparison.Ordinal) < text.LastIndexOf("RepeatedParagraph", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TableCorrelationConsumesEachOccurrenceWhenStoryLocalIndexesRepeat(bool firstBlockHasLogicalOrder) {
        OfficeDocumentModelLocation Location(long? order) => new() { SourceBlockIndex = 0, TableIndex = 0, LogicalOrder = order };
        var source = new OfficeDocumentModel { Blocks = new[] {
            new OfficeDocumentModelBlock { Kind = "table", Text = "FirstFallback", Location = Location(firstBlockHasLogicalOrder ? 0 : null) },
            new OfficeDocumentModelBlock { Text = "BetweenTables", Location = new() { LogicalOrder = 1 } },
            new OfficeDocumentModelBlock { Kind = "table", Text = "SecondFallback", Location = Location(2) }
        }, Tables = new[] {
            new OfficeDocumentModelTable { Columns = new[] { "FirstTable" }, Rows = new[] { new[] { "cell" } }, Location = Location(0) },
            new OfficeDocumentModelTable { Columns = new[] { "SecondTable" }, Rows = new[] { new[] { "cell" } }, Location = Location(2) }
        } };
        string text = PdfReadDocument.Open(source.ToPdfDocumentResult(new PdfProjectionOptions { IncludeMetadata = false }).ToBytes()).ExtractText();
        Assert.DoesNotContain("Fallback", text);
        Assert.True(text.IndexOf("FirstTable", StringComparison.Ordinal) < text.IndexOf("BetweenTables", StringComparison.Ordinal));
        Assert.True(text.IndexOf("BetweenTables", StringComparison.Ordinal) < text.IndexOf("SecondTable", StringComparison.Ordinal));
    }

    [Fact]
    public void PreserveSourcePagesKeepsPhysicalOrderAndHeadingsWithLogicalLocations() {
        var source = new OfficeDocumentModel { Pages = new[] {
            new OfficeDocumentModelPage { Number = 1, Name = "PhysicalFirst", Blocks = new[] {
                new OfficeDocumentModelBlock { Text = "LogicalSecond", Location = new() { LogicalOrder = 1 } } } },
            new OfficeDocumentModelPage { Number = 2, Name = "PhysicalSecond", Blocks = new[] {
                new OfficeDocumentModelBlock { Text = "LogicalFirst", Location = new() { LogicalOrder = 0 } } } }
        } };
        PdfDocumentConversionResult conversion = source.ToPdfDocumentResult(new PdfProjectionOptions { IncludeMetadata = false });
        PdfReadDocument pdf = PdfReadDocument.Open(conversion.ToBytes());
        Assert.Equal(2, pdf.Pages.Count);
        string text = pdf.ExtractText();
        Assert.True(text.IndexOf("LogicalSecond", StringComparison.Ordinal) < text.IndexOf("LogicalFirst", StringComparison.Ordinal));
        Assert.Contains("PhysicalFirst", text); Assert.Contains("PhysicalSecond", text);
        Assert.DoesNotContain(conversion.Warnings, warning => warning.Code == "MODEL_SOURCE_PAGE_LABELS_OMITTED");
    }

    [Fact]
    public void NativeIndentedParagraphAndEveryContinuationStayLiteralAndBounded() {
        byte[] input = Fixture("SampleNewsletter.pub");
        Assert.True(OfficeCompoundFileReader.TryRead(input, out OfficeCompoundFile? compound, out string? error), error);
        const string streamName = "Quill/QuillSub/CONTENTS";
        byte[] quill = (byte[])compound!.Streams[streamName].Clone();
        // Preserve the native text length and every formatting/paragraph offset.
        PublisherDocument original = PublisherDocument.Load(input);
        string prefix = original.TextStories.SelectMany(story => story.Paragraphs)
            .Where(paragraph => paragraph.Label == null)
            .Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text)))
            .First(text => text.Length > 200);
        byte[] needle = Encoding.Unicode.GetBytes(prefix);
        int offset = Find(quill, needle);
        string replacement = "    *value*".PadRight(prefix.Length);
        Encoding.Unicode.GetBytes(replacement).CopyTo(quill, offset);
        input = OfficeCompoundFileWriter.Rewrite(compound, new Dictionary<string, byte[]> { [streamName] = quill });
        OfficeDocumentReadResult result = new OfficeDocumentReaderBuilder().AddPublisherHandler().Build()
            .ReadDocument(input, "indent.pub", new ReaderOptions { MaxChars = 256 });
        OfficeDocumentBlock block = Assert.Single(result.Blocks, item => item.Text.TrimEnd('\n') == replacement);
        Assert.StartsWith("    *value*", block.Text);
        Assert.All(result.Chunks, chunk => Assert.InRange(chunk.Markdown!.Length, 0, 256));
        ReaderChunk[] chunks = result.Chunks.Where(chunk => chunk.Location.BlockAnchor == block.Id).ToArray();
        Assert.True(chunks.Length > 1);
        Assert.All(chunks.Skip(1), chunk => Assert.True(chunk.ContinuesPreviousChunk));
        foreach (ReaderChunk chunk in chunks) {
            string expected = chunk.Text.TrimEnd('\r', '\n');
            if (expected.Length == 0) continue;
            var paragraph = Assert.IsType<OfficeIMO.Markdown.ParagraphBlock>(Assert.Single(MarkdownReader.Parse(chunk.Markdown!).Blocks));
            var literal = new StringBuilder();
            ((IPlainTextMarkdownInline)paragraph.Inlines).AppendPlainText(literal);
            Assert.Equal(expected, literal.ToString());
        }
    }

    private static int Find(byte[] input, byte[] needle) {
        for (int index = 0; index <= input.Length - needle.Length; index++)
            if (input.AsSpan(index, needle.Length).SequenceEqual(needle)) return index;
        throw new InvalidOperationException("The fixture's native paragraph was not found in its Quill stream.");
    }
    private static byte[] Fixture(string name) => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "PublisherFixtures", name));
    private static string Normalize(string text) => Regex.Replace(text, @"\s+", " ").Trim();
}
