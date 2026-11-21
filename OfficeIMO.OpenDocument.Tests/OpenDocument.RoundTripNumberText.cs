using OfficeIMO.OpenDocument;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class OpenDocumentRoundTripNumberTextTests {
    // A cell's office:value uses the shortest round-trip digits; .NET Framework's "R" wrote this value with 17 digits.
    [Fact]
    public void CellValueUsesTheShortestRoundTripDigits() {
        OdsDocument document = OdsDocument.Create();
        OdsCell cell = document.AddSheet("Data").Cell(0, 0);
        cell.SetNumber(0.4694715627858908D);

        Assert.Equal("0.4694715627858908", (string?)cell.Element.Attribute(OdfNamespaces.Office + "value"));
    }
}
