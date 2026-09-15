using System.Text;
using OfficeIMO.ContentSafety;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class SvgContentSafetyCssBoundaryTests {
    [Fact]
    public void InvalidXmlSpaceCasingFailsClosed() {
        byte[] svg = Encoding.UTF8.GetBytes(
            "<svg xmlns='http://www.w3.org/2000/svg' xml:space='PRESERVE' width='220' height='120' viewBox='0 0 220 120'>" +
            "<rect width='220' height='120' fill='white'/>" +
            "<text font-family='OfficeIMO Shaping Test' font-size='20' x='10' y='35'>      A</text>" +
            "<rect width='40' height='60' fill='white'/></svg>");

        Assert.Throws<InvalidDataException>(() => OfficeSvgDrawingReader.InspectContentSafety(svg));
    }

    [Fact]
    public void InvalidNestedCustomPropertyUsesTheCurrentVarFallback() {
        byte[] svg = Svg(
            "<text style='--visibility:var(--missing);display:var(--visibility,none)' x='10' y='35'>nested invalid custom fallback</text>");

        OfficeContentSafetyFinding finding = Assert.Single(
            OfficeSvgDrawingReader.InspectContentSafety(svg).Findings,
            item => item.TextPreview == "nested invalid custom fallback");

        Assert.Equal(OfficeContentConcealmentKind.HiddenByProperty, finding.Kind);
        Assert.Equal(OfficeContentCleanupCapability.RemoveText, finding.CleanupCapability);
    }

    private static byte[] Svg(string body) => Encoding.UTF8.GetBytes(
        "<svg xmlns='http://www.w3.org/2000/svg' width='220' height='120' viewBox='0 0 220 120'>" + body + "</svg>");
}
