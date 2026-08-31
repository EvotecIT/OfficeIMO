using System.Collections.Generic;
using System.Xml.Linq;

namespace OfficeIMO.Bibliography.Tests;

public sealed class BibliographyReviewWave60RegressionTests {
    [Fact]
    public void EndNote_retained_element_serialization_stops_at_the_value_limit_before_allocating_the_full_XML() {
        XElement element = new XElement("extra", Enumerable.Range(0, 250_000).Select(static _ => new XElement("x")));
        var items = new List<BibliographyItem>();
        var limits = new BibliographyLimitGuard(new BibliographyReadOptions { MaximumValueLength = 1024 });
#if NET472
        Assert.Throws<BibliographyLimitException>(() => { EndNoteXmlCodec.SerializeBoundedElement(element, items, limits); });
#else
        long before = GC.GetAllocatedBytesForCurrentThread();
        Assert.Throws<BibliographyLimitException>(() => { EndNoteXmlCodec.SerializeBoundedElement(element, items, limits); });
        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;

        Assert.True(allocated < 2 * 1024 * 1024, $"Bounded EndNote serialization allocated {allocated:N0} bytes before rejecting the element.");
#endif
    }

    [Fact]
    public void EndNote_retained_element_serialization_observes_cancellation_between_nodes() {
        XElement element = new XElement("extra", Enumerable.Range(0, 500_000).Select(static _ => new XElement("x")));
        var items = new List<BibliographyItem>();
        var limits = new BibliographyLimitGuard(new BibliographyReadOptions { MaximumValueLength = 8 * 1024 * 1024 });

        BibliographyCancellationTest.AssertObserved(token =>
            EndNoteXmlCodec.SerializeBoundedElement(element, items, limits, token));
    }

    [Fact]
    public void CSL_retained_escaped_native_JSON_observes_cancellation_during_canonical_write() {
        string escapedValue = string.Concat(Enumerable.Repeat("\\u0061", 4 * 1024 * 1024));
        string source = "[{\"id\":\"x\",\"type\":\"book\",\"custom\":\"" + escapedValue + "\",\"title\":\"Before\"}]";
        var options = new BibliographyReadOptions { MaximumInputCharacters = source.Length + 1, MaximumValueLength = 4 * 1024 * 1024 + 1 };
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.CslJson, options).Document;
        document.Items[0].Title = "After";

        BibliographyCancellationTest.AssertObserved(token =>
            document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical }, token));
    }

    [Fact]
    public void CSL_strict_native_JSON_keeps_its_exact_raw_formatting_after_segmented_validation() {
        const string source = "[{\"id\":\"x\",\"type\":\"book\",\"custom\":{ \"a\" : 1 },\"title\":\"Before\"}]";
        BibliographyDocument document = BibliographyDocument.Parse(source, BibliographyFormat.CslJson).Document;
        document.Items[0].Title = "After";

        BibliographyWriteResult written = document.Write(new BibliographyWriteOptions { Mode = BibliographyWriterMode.Canonical, RequireNoLoss = true });

        Assert.Contains("{ \"a\" : 1 }", written.Content, StringComparison.Ordinal);
        Assert.Equal("{ \"a\" : 1 }", Assert.Single(BibliographyDocument.Parse(written.Content, BibliographyFormat.CslJson).Document.Items[0].NativeFields).RawValue);
    }
}
