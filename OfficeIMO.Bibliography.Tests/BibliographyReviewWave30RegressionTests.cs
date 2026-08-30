using System.Collections.Generic;
using System.Xml.Linq;

namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewWave30RegressionTests {
    [Fact]
    public void Signed_CSL_years_survive_unrelated_strict_edits() {
        const string source = "[{\"id\":\"x\",\"type\":\"book\",\"title\":\"Before\",\"issued\":{\"date-parts\":[[-44,2,3],[-43,4,5]]}}]";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.CslJson).Document;
        document.Items[0].Title = "After";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyDate reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.CslJson).Document.Items[0].Dates);

        Assert.Equal(-44, reopened.Year);
        Assert.Equal(2, reopened.Month);
        Assert.Equal(3, reopened.Day);
        Assert.Equal(-43, reopened.EndYear);
        Assert.Equal(4, reopened.EndMonth);
        Assert.Equal(5, reopened.EndDay);
        Assert.DoesNotContain(written.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV218");
    }

    [Fact]
    public void CSL_numeric_parts_without_a_year_remain_lossy() {
        var document = new BibliographyDocument(BibliographyFormat.CslJson);
        var item = new BibliographyItem { Key = "x", Type = BibliographyItemType.Book };
        item.Dates.Add(new BibliographyDate { Role = BibliographyDateRole.Issued, Month = 2 });
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV218" && diagnostic.Field == "dates.Issued");
    }

    [Fact]
    public void Generated_byte_encoding_observes_cancellation_between_chunks() {
        using var cancellation = new CancellationTokenSource();
        var encoding = new CancelAfterFirstConvertEncoding(cancellation);

        Assert.Throws<OperationCanceledException>(() =>
            BibliographyEncoding.Encode(new string('x', 16 * 1024), encoding, cancellation.Token));
    }

    [Fact]
    public void Generated_output_encoding_inspection_observes_cancellation_between_chunks() {
        using var cancellation = new CancellationTokenSource();
        var document = new BibliographyDocument(BibliographyFormat.CslJson);
        document.Items.Add(new BibliographyItem { Key = "x", Type = BibliographyItemType.Book, Title = new string('x', 16 * 1024) });

        Assert.Throws<OperationCanceledException>(() => document.Write(new BibliographyWriteOptions {
            Encoding = new CancelAfterFirstConvertEncoding(cancellation)
        }, cancellation.Token));
    }

    [Theory]
    [InlineData(BibliographyFormat.BibTex)]
    [InlineData(BibliographyFormat.BibLatex)]
    public void Bib_identifier_scheme_casing_reopens_exactly(BibliographyFormat format) {
        var document = new BibliographyDocument(format);
        var item = new BibliographyItem { Key = "x", Type = BibliographyItemType.Book, Title = "Identifiers" };
        item.Identifiers.Add(new BibliographyIdentifier("DOI", "10.1000/example"));
        item.Identifiers.Add(new BibliographyIdentifier("IsBn", "978-1-4028-9462-6"));
        item.Identifiers.Add(new BibliographyIdentifier("PMID", "123"));
        document.Items.Add(item);

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, format).Document.Items);

        Assert.Equal(new[] { "DOI", "IsBn", "PMID" }, reopened.Identifiers.Select(static identifier => identifier.Scheme));
        Assert.Contains("DOI = {10.1000/example}", written.Content, StringComparison.Ordinal);
        Assert.Contains("IsBn = {978-1-4028-9462-6}", written.Content, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("library", "records")]
    [InlineData("Library", "Records")]
    public void EndNote_root_and_records_element_names_survive_strict_edits(string rootName, string recordsName) {
        string source = "<" + rootName + "><" + recordsName + "><record><rec-number>1</rec-number><ref-type name=\"Book\">6</ref-type><titles><title>Before</title></titles></record></" + recordsName + "></" + rootName + ">";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;
        document.Items[0].Title = "After";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyDocument reopened = BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document;

        Assert.Contains("<" + rootName + ">", written.Content, StringComparison.Ordinal);
        Assert.Contains("<" + recordsName + ">", written.Content, StringComparison.Ordinal);
        Assert.Equal("After", Assert.Single(reopened.Items).Title);
    }

    [Theory]
    [InlineData(BibliographyFormat.Ris, "given")]
    [InlineData(BibliographyFormat.Ris, "family")]
    [InlineData(BibliographyFormat.Ris, "suffix")]
    [InlineData(BibliographyFormat.Nbib, "given")]
    [InlineData(BibliographyFormat.Nbib, "family")]
    [InlineData(BibliographyFormat.Nbib, "suffix")]
    [InlineData(BibliographyFormat.EndNoteXml, "given")]
    [InlineData(BibliographyFormat.EndNoteXml, "family")]
    [InlineData(BibliographyFormat.EndNoteXml, "suffix")]
    public void Tagged_and_EndNote_structured_name_whitespace_is_diagnosed(BibliographyFormat format, string component) {
        var document = new BibliographyDocument(format);
        var item = new BibliographyItem { Key = "1", Type = BibliographyItemType.Book };
        var name = new BibliographyName { Given = "John", Family = "Smith", Suffix = "Jr." };
        if (component == "given") name.Given = " John ";
        else if (component == "family") name.Family = " Smith ";
        else name.Suffix = " Jr. ";
        item.Contributors.Add(new BibliographyContributor(BibliographyContributorRole.Author, name));
        if (format == BibliographyFormat.Nbib) item.Identifiers.Add(new BibliographyIdentifier("PMID", "1"));
        document.Items.Add(item);

        BibliographyConversionLossException exception = Assert.Throws<BibliographyConversionLossException>(() =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true }));

        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Code == "BIBCONV243" && diagnostic.Field == "contributors");
    }

    [Fact]
    public void EndNote_record_projection_observes_cancellation_after_materialization() {
        var record = new XElement("record", new XElement("rec-number", "1"), new XElement("ref-type", new XAttribute("name", "Book"), "6"));
        for (int index = 0; index < 200_000; index++) record.Add(new XElement("custom", index.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        var options = new BibliographyReadOptions { MaximumValueCount = 500_000 };
        var items = new List<BibliographyItem>();
        var diagnostics = new List<BibliographyDiagnostic>();
        var guard = new BibliographyDiagnosticGuard(options, diagnostics, items);
        var limits = new BibliographyLimitGuard(options);
        using var cancellation = new CancellationTokenSource();
        using var started = new ManualResetEventSlim();
        var cancelThread = new Thread(() => { started.Wait(); Thread.Sleep(1); cancellation.Cancel(); });
        cancelThread.Start();

        try {
            started.Set();
            Assert.Throws<OperationCanceledException>(() => EndNoteXmlCodec.ParseRecord(record, items, limits, guard, cancellation.Token));
        } finally {
            cancelThread.Join();
        }
    }

    [Fact]
    public void Redundant_EndNote_periodical_titles_survive_unrelated_strict_edits() {
        const string source = "<xml><records><record><rec-number>1</rec-number><ref-type name=\"Journal Article\">17</ref-type><titles><title>Before</title><secondary-title>Journal</secondary-title></titles><periodical><full-title>Journal</full-title></periodical></record></records></xml>";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.EndNoteXml).Document;
        document.Items[0].Title = "After";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });
        BibliographyItem reopened = Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.EndNoteXml).Document.Items);

        Assert.Contains("<secondary-title>Journal</secondary-title>", written.Content, StringComparison.Ordinal);
        Assert.Contains("<periodical>", written.Content, StringComparison.Ordinal);
        Assert.Contains(reopened.NativeFields, field => field.Format == BibliographyFormat.EndNoteXml && field.Name == "periodical" && field.Value.Trim() == "Journal");
    }

    private sealed class CancelAfterFirstConvertEncoding : Encoding {
        private readonly Encoding _inner = new UTF8Encoding(false);
        private readonly CancellationTokenSource _cancellation;

        internal CancelAfterFirstConvertEncoding(CancellationTokenSource cancellation) => _cancellation = cancellation;

        public override int GetByteCount(char[] chars, int index, int count) => _inner.GetByteCount(chars, index, count);
        public override int GetBytes(char[] chars, int charIndex, int charCount, byte[] bytes, int byteIndex) => _inner.GetBytes(chars, charIndex, charCount, bytes, byteIndex);
        public override int GetCharCount(byte[] bytes, int index, int count) => _inner.GetCharCount(bytes, index, count);
        public override int GetChars(byte[] bytes, int byteIndex, int byteCount, char[] chars, int charIndex) => _inner.GetChars(bytes, byteIndex, byteCount, chars, charIndex);
        public override int GetMaxByteCount(int charCount) => _inner.GetMaxByteCount(charCount);
        public override int GetMaxCharCount(int byteCount) => _inner.GetMaxCharCount(byteCount);
        public override byte[] GetPreamble() => _inner.GetPreamble();
        public override Decoder GetDecoder() => _inner.GetDecoder();
        public override Encoder GetEncoder() => new CancelAfterFirstConvertEncoder(_inner.GetEncoder(), _cancellation);
    }

    private sealed class CancelAfterFirstConvertEncoder : Encoder {
        private readonly Encoder _inner;
        private readonly CancellationTokenSource _cancellation;

        internal CancelAfterFirstConvertEncoder(Encoder inner, CancellationTokenSource cancellation) {
            _inner = inner;
            _cancellation = cancellation;
        }

        public override int GetByteCount(char[] chars, int index, int count, bool flush) => _inner.GetByteCount(chars, index, count, flush);
        public override int GetBytes(char[] chars, int charIndex, int charCount, byte[] bytes, int byteIndex, bool flush) => _inner.GetBytes(chars, charIndex, charCount, bytes, byteIndex, flush);

        public override void Convert(char[] chars, int charIndex, int charCount, byte[] bytes, int byteIndex, int byteCount, bool flush, out int charsUsed, out int bytesUsed, out bool completed) {
            _inner.Convert(chars, charIndex, charCount, bytes, byteIndex, byteCount, flush, out charsUsed, out bytesUsed, out completed);
            _cancellation.Cancel();
        }
    }
}
