namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewWave21RegressionTests {
    [Theory]
    [InlineData("parse-explicit", false)]
    [InlineData("parse-detected", false)]
    [InlineData("stream", false)]
    [InlineData("stream", true)]
    [InlineData("path", false)]
    [InlineData("path", true)]
    [InlineData("path-small", false)]
    [InlineData("path-small", true)]
    public async Task Stream_load_routes_enforce_character_limits_while_decoding(string route, bool asynchronous) {
        const string smallDetectedSource = "TY  - BOOK\nID  - x\nTI  - Bounded\nER  -\n";
        string text = route == "parse-explicit" ? "@book{x,title={too long}}"
            : route == "parse-detected" || route == "path-small" ? smallDetectedSource
            : "@book{x,title={" + new string('x', 128 * 1024) + "}}";
        byte[] source = Encoding.UTF8.GetBytes(text);
        var options = new BibliographyReadOptions {
            MaximumInputBytes = route == "parse-detected" || route == "path-small" ? 1024 : source.Length + 1,
            MaximumInputCharacters = route.StartsWith("parse", StringComparison.Ordinal) || route == "path-small" ? 8 : 32
        };
        if (route == "parse-explicit") {
            Assert.Throws<InvalidDataException>(() => BibliographyDocument.Parse(text, BibliographyFormat.BibTex, options));
            return;
        }
        if (route == "parse-detected") {
            Assert.Throws<InvalidDataException>(() => BibliographyDocument.Parse(text, options));
            return;
        }
        string? path = route.StartsWith("path", StringComparison.Ordinal) ? Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".unknown") : null;
        Encoding? pathEncoding = route == "path-small" ? null : Encoding.UTF8;
        if (path != null) File.WriteAllBytes(path, source);
        try {
            if (asynchronous) {
                if (path != null) await Assert.ThrowsAsync<InvalidDataException>(() => BibliographyDocument.LoadAsync(path, options: options, encoding: pathEncoding));
                else {
                    using var stream = new MemoryStream(source);
                    await Assert.ThrowsAsync<InvalidDataException>(() => BibliographyDocument.LoadAsync(stream, BibliographyFormat.BibLatex, options, Encoding.UTF8));
                }
            } else if (path != null) Assert.Throws<InvalidDataException>(() => BibliographyDocument.Load(path, options: options, encoding: pathEncoding));
            else {
                using var stream = new MemoryStream(source);
                Assert.Throws<InvalidDataException>(() => BibliographyDocument.Load(stream, BibliographyFormat.BibLatex, options, Encoding.UTF8));
            }
        } finally {
            if (path != null) File.Delete(path);
        }
    }

    [Fact]
    public void Escaped_CSL_native_strings_use_decoded_value_limits() {
        const string source = "[{\"id\":\"i\",\"type\":\"book\",\"x\":\"\\u0061\\u0061\"}]";
        var options = new BibliographyReadOptions { MaximumValueLength = 4 };

        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.CslJson, options);
        BibliographyNativeField field = Assert.Single(Assert.Single(read.Document.Items).NativeFields);

        Assert.Equal("aa", field.Value);
        Assert.Equal("\"\\u0061\\u0061\"", field.RawValue);
        Assert.DoesNotContain(read.Diagnostics, diagnostic => diagnostic.Code == "BIBLIM001");
    }

    [Theory]
    [InlineData("")]
    [InlineData(" ")]
    public void Blank_EndNote_notes_remain_distinct_from_an_absent_note(string note) {
        var document = new BibliographyDocument(BibliographyFormat.EndNoteXml);
        var item = new BibliographyItem { Key = "1", Type = BibliographyItemType.Book, Title = "Note" };
        item.Notes.Add(note);
        document.Items.Add(item);

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document.Items);

        Assert.Equal(note, Assert.Single(reopened.Notes));
    }

    [Fact]
    public void Conversion_inspection_observes_cancellation() {
        var document = new BibliographyDocument(BibliographyFormat.CslJson);
        document.Items.Add(new BibliographyItem { Key = "x", Type = BibliographyItemType.Book, Title = "Canceled" });
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() =>
            BibliographyConversionInspector.Inspect(document, BibliographyFormat.CslJson, new BibliographyConversionReport(), cancellation.Token));
    }

    [Fact]
    [Trait("Category", "ResourcePerformanceEvidence")]
    public void Tagged_aggregate_limits_apply_before_continuation_concatenation() {
#if BIBLIOGRAPHY_PERFORMANCE_EVIDENCE
        const int length = 512 * 1024;
#else
        const int length = 128;
#endif
        string value = new string('x', length);
        string source = "TY  - JOUR\nTI  - " + value + "\n      " + value + "\nER  -";
        var options = new BibliographyReadOptions { MaximumValueLength = value.Length + 1 };
#if !BIBLIOGRAPHY_PERFORMANCE_EVIDENCE || NET472
        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.Ris, options);
#else
        long before = GC.GetAllocatedBytesForCurrentThread();
        BibliographyReadResult read = BibliographyDocument.Parse(source, BibliographyFormat.Ris, options);
        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;
#endif

        Assert.Contains(read.Diagnostics, diagnostic => diagnostic.Code == "BIBLIM001");
#if BIBLIOGRAPHY_PERFORMANCE_EVIDENCE && !NET472
        Assert.True(allocated < 5 * 1024 * 1024, $"Oversized tagged continuation allocated {allocated:N0} bytes before rejection.");
#endif
    }

    [Fact]
    public void Continued_NBIB_publication_types_rebind_the_typed_item() {
        const string source = "PMID- 1\nPT  - Book\n      Chapter\nTI  - Continued type";

        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.Nbib).Document;
        Assert.Equal(BibliographyItemType.Chapter, Assert.Single(document.Items).Type);

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.Nbib).Document.Items);
        Assert.Equal(BibliographyItemType.Chapter, reopened.Type);
        Assert.Contains(reopened.NativeFields, field => field.Name == "PT" && field.Value == "Book Chapter");
    }

    [Fact]
    public void BibLaTeX_preserves_supported_date_role_order() {
        var document = new BibliographyDocument(BibliographyFormat.BibLatex);
        var item = new BibliographyItem { Key = "x", Type = BibliographyItemType.Book, Title = "Dates" };
        item.Dates.Add(new BibliographyDate { Role = BibliographyDateRole.Accessed, Year = 2026, Month = 8, Day = 30 });
        item.Dates.Add(new BibliographyDate { Role = BibliographyDateRole.Issued, Year = 2020 });
        document.Items.Add(item);

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.BibLatex).Document.Items);

        Assert.Equal(new[] { BibliographyDateRole.Accessed, BibliographyDateRole.Issued }, reopened.Dates.Select(date => date.Role));
    }
}
