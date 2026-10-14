using System.Text;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Visio.Tests;

public sealed class VisioDrawingGeometryTests {
    private static readonly XNamespace Legacy = "http://schemas.microsoft.com/visio/2003/core";

    [Theory]
    [InlineData("control", true, true)]
    [InlineData("NoShow", false, false)]
    [InlineData("NoLine", true, false)]
    [InlineData("NoFill", false, true)]
    public void NativeFillAndOutlineVisibilityRemainIndependentInPageScenes(string flag, bool fill, bool stroke) {
        VisioDocument source = Load(ShapeXml(flag));
        foreach (VisioDocument document in Reopened(source)) {
            byte[] before = document.ToLegacyXmlResult().Value;
            OfficeDrawing scene = document.ToDrawings().Value[0];
            Assert.Equal(fill, scene.Shapes.Any(item => item.Shape.FillColor?.R == 255));
            Assert.Equal(stroke, scene.Shapes.Any(item => item.Shape.StrokeColor?.B == 255));
            Assert.Equal(before, document.ToLegacyXmlResult().Value);
        }
    }

    [Fact]
    public void MultipleNativeConnectorOutlinesAreKeptAndUnprojectedFillAndOpaqueContentAreReported() {
        VisioDocument source = Load(ConnectorXml());
        foreach (VisioDocument document in Reopened(source)) {
            byte[] before = document.ToLegacyXmlResult().Value;
            var result = document.ToDrawings();
            Assert.Equal(2, result.Value[0].Shapes.Count(item => item.Shape.StrokeColor?.B == 255));
            Assert.Contains(result.Report.FidelityDiagnostics, item => item.Code == "VISIO_DRAWING_CONNECTOR_FILL" && item.LossKind == OfficeConversionLossKind.Omission);
            Assert.Contains(result.Report.FidelityDiagnostics, item => item.Code == "VISIO_DRAWING_CONNECTOR_CONTENT" && item.LossKind == OfficeConversionLossKind.Omission);
            Assert.Contains(result.Report.FidelityDiagnostics, item => item.Code == "VISIO_DRAWING_CONNECTOR_ROUTE" && item.LossKind == OfficeConversionLossKind.Approximation);
            Assert.Throws<OfficeConversionException>(() => result.Report.RequireNoLoss());
            Assert.Equal(before, document.ToLegacyXmlResult().Value);
        }
    }

    private static IEnumerable<VisioDocument> Reopened(VisioDocument source) {
        yield return source;
        yield return VisioDocument.LoadLegacyXml(new MemoryStream(source.ToLegacyXmlResult().Value)).Value;
        yield return VisioDocument.Load(new MemoryStream(source.ToBytes()));
    }

    private static VisioDocument Load(string xml) => VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(xml))).Value;

    private static string ShapeXml(string flag) => $"""
        <VisioDocument xmlns="{Legacy}"><Pages><Page ID="0" Name="Geometry"><PageSheet><PageProps><PageWidth>4</PageWidth><PageHeight>3</PageHeight></PageProps></PageSheet>
        <Shapes><Shape ID="1" NameU="Controlled" Type="Shape"><XForm><PinX>2</PinX><PinY>1.5</PinY><Width>2</Width><Height>1</Height><LocPinX>1</LocPinX><LocPinY>0.5</LocPinY><Angle>0</Angle></XForm>
        <Line><LineWeight>0.1</LineWeight><LineColor>#0000FF</LineColor><LinePattern>1</LinePattern></Line><Fill><FillForegnd>#FF0000</FillForegnd><FillPattern>1</FillPattern></Fill>
        <Geom IX="0"><NoFill>{(flag == "NoFill" ? 1 : 0)}</NoFill><NoLine>{(flag == "NoLine" ? 1 : 0)}</NoLine><NoShow>{(flag == "NoShow" ? 1 : 0)}</NoShow>
        <MoveTo IX="0"><X>0</X><Y>0</Y></MoveTo><LineTo IX="1"><X>2</X><Y>0</Y></LineTo><LineTo IX="2"><X>2</X><Y>1</Y></LineTo><LineTo IX="3"><X>0</X><Y>1</Y></LineTo><LineTo IX="4"><X>0</X><Y>0</Y></LineTo></Geom>
        </Shape></Shapes></Page></Pages></VisioDocument>
        """;

    private static string ConnectorXml() => $"""
        <VisioDocument xmlns="{Legacy}"><Pages><Page ID="0" Name="Connector"><PageSheet><PageProps><PageWidth>4</PageWidth><PageHeight>3</PageHeight></PageProps></PageSheet><Shapes>
        <Shape ID="1" NameU="Connector" Type="Shape"><XForm><PinX>2</PinX><PinY>1.5</PinY><Width>3</Width><Height>1</Height><LocPinX>1.5</LocPinX><LocPinY>0.5</LocPinY><Angle>0</Angle></XForm>
        <XForm1D><BeginX>0.5</BeginX><BeginY>1</BeginY><EndX>3.5</EndX><EndY>1</EndY></XForm1D><Line><LineWeight>0.1</LineWeight><LineColor>#0000FF</LineColor><LinePattern>1</LinePattern></Line>
        <Geom IX="0"><NoFill>1</NoFill><MoveTo IX="0"><X>0</X><Y>0</Y></MoveTo><LineTo IX="1"><X>3</X><Y>0</Y></LineTo></Geom>
        <Geom IX="1"><NoFill>0</NoFill><MoveTo IX="0"><X>0</X><Y>1</Y></MoveTo><LineTo IX="1"><X>3</X><Y>1</Y></LineTo><LineTo IX="2"><X>2</X><Y>0.5</Y></LineTo></Geom>
        <Custom xmlns="urn:officeimo:test">opaque</Custom></Shape></Shapes></Page></Pages></VisioDocument>
        """;
}
