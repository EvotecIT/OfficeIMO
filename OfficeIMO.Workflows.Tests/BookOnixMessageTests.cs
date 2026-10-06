using OfficeIMO.Epub;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class BookOnixMessageTests {
    [Fact]
    public void ProductsKeepOrderExactRecordContentAndPublicationEvidence() {
        var first = Export("one", "9780306406157");
        var second = Export("two", "9781861972712");
        var inputs = new[] { second, first };
        byte[] before = first.Bytes.ToArray();
        var message = BookOnixMessage.Create(inputs, BookOnixTests.TestSchema());
        inputs[0] = first;
        Assert.Same(second, message.Products[0]);
        Assert.Same(first, message.Products[1]);
        Assert.Equal(before, first.Bytes);
        XNamespace ns = BookProject.OnixNamespace;
        var xml = XDocument.Load(new MemoryStream(message.Bytes));
        Assert.Single(xml.Root!.Elements(ns + "Header"));
        Assert.Equal(new[] { "two", "one" }, xml.Descendants(ns + "RecordReference").Select(e => e.Value));
        var records = xml.Root.Elements(ns + "Product").ToArray();
        Assert.True(XNode.DeepEquals(XDocument.Load(new MemoryStream(second.Bytes)).Root!.Element(ns + "Product"), records[0]));
        Assert.True(XNode.DeepEquals(XDocument.Load(new MemoryStream(first.Bytes)).Root!.Element(ns + "Product"), records[1]));
        Assert.Equal(message.Bytes, BookOnixMessage.Create([second, first], BookOnixTests.TestSchema()).Bytes);
        Assert.Equal(first.Bytes, BookOnixMessage.Create([first], BookOnixTests.TestSchema()).Bytes);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void RepeatedRecordReferencesOrIsbnsAreRejected(bool sameReference) {
        var first = Export("one", "9780306406157");
        var second = Export(sameReference ? "one" : "two", sameReference ? "9781861972712" : "9780306406157");
        Assert.Throws<ArgumentException>(() => BookOnixMessage.Create([first, second], BookOnixTests.TestSchema()));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void MixedSendersOrTimestampsAreNotSilentlyRewritten(bool sender) {
        var first = Export("one", "9780306406157");
        var options = BookOnixTests.Options() with { RecordReference = "two" };
        options = sender ? options with { SenderName = "Another Press" } : options with { SentAt = options.SentAt.AddSeconds(1) };
        var second = Project("9781861972712").ExportOnix(options, BookOnixTests.TestSchema());
        Assert.Throws<ArgumentException>(() => BookOnixMessage.Create([first, second], BookOnixTests.TestSchema()));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void MutatedXmlOrEpubCannotRetainTheOriginalAssociation(bool xml) {
        var product = Export("one", "9780306406157");
        byte[] bytes = xml ? product.Bytes : product.Publication.Bytes;
        bytes[bytes.Length / 2] ^= 1;
        Assert.Throws<InvalidDataException>(() => BookOnixMessage.Create([product], BookOnixTests.TestSchema()));
    }

    [Fact]
    public void SchemaCancellationAndProductBoundsAreEnforced() {
        var product = Export("one", "9780306406157");
        Assert.Throws<ArgumentException>(() => BookOnixMessage.Create([], BookOnixTests.TestSchema()));
        Assert.Throws<ArgumentException>(() => BookOnixMessage.Create(Enumerable.Repeat(product, 1001).ToArray(), BookOnixTests.TestSchema()));
        Assert.Throws<ArgumentException>(() => BookOnixMessage.Create([product], new()));
        Assert.Throws<InvalidDataException>(() => BookOnixMessage.Create([product], BookOnixTests.TestSchema(rejectChildren: true)));
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => BookOnixMessage.Create([product], BookOnixTests.TestSchema(), cancellation.Token));
    }

    [Fact]
    public void CombinedXmlHasAnAggregateBound() {
        var records = new List<BookOnixExportResult>();
        for (int index = 0; index < 18; index++) {
            string prefix = "97800000" + index.ToString("D4");
            int sum = prefix.Select((digit, position) => (digit - '0') * (position % 2 == 0 ? 1 : 3)).Sum();
            string isbn = prefix + ((10 - sum % 10) % 10);
            records.Add(Project(isbn).ExportOnix(BookOnixTests.Options() with {
                RecordReference = "edition-" + index,
                Contributors = Enumerable.Repeat(new BookOnixContributor(new string('&', 2000), BookOnixContributorRole.Author), 100).ToArray()
            }, BookOnixTests.TestSchema()));
        }
        Assert.True(records.Sum(record => record.Bytes.LongLength) > 16L * 1024 * 1024);
        Assert.Throws<InvalidDataException>(() => BookOnixMessage.Create(records, BookOnixTests.TestSchema()));
    }

    private static BookOnixExportResult Export(string reference, string isbn) => Project(isbn).ExportOnix(
        BookOnixTests.Options() with { RecordReference = reference }, BookOnixTests.TestSchema());

    private static BookProject Project(string isbn) {
        var project = BookProject.Create("Book " + isbn);
        project.Publication.AddIdentifier("isbn", new EpubIdentifierMetadata { Value = isbn, Kind = EpubIdentifierKind.Isbn13 });
        return project;
    }
}
