using System.IO;
using System.Text;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public class VisioTextSizeTests {
    [Theory]
    [InlineData("", "0.25", 18)]
    [InlineData(" Unit='PT'", "4", 288)]
    public void NativeSizeIsAnInternalLengthRegardlessOfDisplayUnit(string unit, string value, double points) {
        string xml = "<VisioDocument xmlns='http://schemas.microsoft.com/visio/2003/core'><Pages><Page ID='0'><Shapes><Shape ID='1'><Char IX='0'><Size" + unit + ">" + value + "</Size></Char><Text>label</Text></Shape></Shapes></Page></Pages></VisioDocument>";
        var document = VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(xml))).Value;
        Assert.Equal(points, document.Pages[0].Shapes[0].TextStyle!.Size);
        Assert.Equal(points, VisioDocument.Load(new MemoryStream(document.ToBytes())).Pages[0].Shapes[0].TextStyle!.Size);
    }

    [Fact]
    public void LargeShapeAndConnectorTextSizesSurvivePackageRoundTrip() {
        var document = VisioDocument.Create(); var page = document.AddPage("Page");
        var first = page.AddRectangle(1, 1, 1, 1, "large");
        first.TextStyle = new VisioTextStyle { Size = 288 };
        page.AddConnector(first, page.AddRectangle(3, 1, 1, 1)).TextStyle = new VisioTextStyle { Size = 360 };
        var loaded = VisioDocument.Load(new MemoryStream(document.ToBytes())).Pages[0];
        Assert.Equal(288, loaded.Shapes[0].TextStyle!.Size);
        Assert.Equal(360, loaded.Connectors[0].TextStyle!.Size);
    }
}
