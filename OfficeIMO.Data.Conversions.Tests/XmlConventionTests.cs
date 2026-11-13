using System.Xml.Linq;
using OfficeIMO.Data.Conversions;
using Xunit;

namespace OfficeIMO.Data.Conversions.Tests;

// The documented XML convention, both ways: data XML converts to the expected JSON and back to an equal document.
public class XmlConventionTests {
    private const string Orders = """
        <?xml version="1.0" encoding="utf-8"?>
        <orders xmlns:x="urn:x" version="2">
          <order id="1"><item>Pen</item><item>Ink</item><x:note>rush</x:note><gift xsi:nil="true" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"/></order>
          <order id="2"><item>Pad</item><empty/></order>
        </orders>
        """;

    [Fact]
    public void DataXmlFollowsTheConvention() {
        (string json, _) = Convert.Run(Orders, DataFormat.Xml, DataFormat.Json, new DataConversionOptions {
            JsonWriter = new System.Text.Json.JsonWriterOptions()
        });
        Assert.Equal(
            "{\"orders\":{\"@xmlns:x\":\"urn:x\",\"@version\":\"2\",\"order\":[" +
            "{\"@id\":\"1\",\"item\":[\"Pen\",\"Ink\"],\"x:note\":\"rush\",\"gift\":null}," +
            "{\"@id\":\"2\",\"item\":\"Pad\",\"empty\":\"\"}]}}\n", json);
    }

    [Fact]
    public void DataXmlRoundTripsThroughJson() {
        (string json, _) = Convert.Run(Orders, DataFormat.Xml, DataFormat.Json);
        (string xml, OfficeConversionReport report) = Convert.Run(json, DataFormat.Json, DataFormat.Xml);
        Assert.DoesNotContain(OfficeDiagnosticCodes.RootWrapped, Convert.Codes(report));
        Assert.True(XNode.DeepEquals(Normalize(XDocument.Parse(Orders)), Normalize(XDocument.Parse(xml))), xml);
    }

    // xsi declarations move to where nil is written; compare elements, attributes and text only.
    private static XElement Normalize(XDocument document) {
        XElement root = new(document.Root!);
        foreach (XAttribute attribute in root.DescendantsAndSelf().Attributes().Where(a => a.IsNamespaceDeclaration && a.Value.Contains("XMLSchema-instance")).ToList()) {
            attribute.Remove();
        }

        foreach (XElement element in root.DescendantsAndSelf().Where(e => e.IsEmpty)) element.Value = string.Empty;
        return root;
    }
}
