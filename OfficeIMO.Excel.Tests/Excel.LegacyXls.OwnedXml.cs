using OfficeIMO.Excel.LegacyXls.Write;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Fact]
    public void LegacyXls_StreamingStylesIgnoreExtensionAndAlternateContentElements() {
        const string xml = """
            <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                        xmlns:custom="urn:test">
              <numFmts>
                <numFmt numFmtId="164" formatCode="yyyy-mm-dd hh:mm" />
                <numFmt numFmtId="165" formatCode="[h]:mm:ss" />
              </numFmts>
              <fonts><font><sz val="11" /><color theme="1" /><name val="Aptos" /><family val="2" /></font></fonts>
              <cellXfs><xf numFmtId="0" /><xf numFmtId="164" /><xf numFmtId="165" /><xf numFmtId="14" /><xf numFmtId="46" /></cellXfs>
            </styleSheet>
            """;
        string decorated = xml
            .Replace("<numFmts>", """
                <numFmts>
                  <extLst><ext uri="urn:test"><numFmt numFmtId="164" formatCode="00000" /></ext></extLst>
                  <mc:AlternateContent>
                    <mc:Choice Requires="custom"><numFmt numFmtId="165" formatCode="0.00" /></mc:Choice>
                    <mc:Fallback><numFmt numFmtId="165" formatCode="0.0" /></mc:Fallback>
                  </mc:AlternateContent>
                """)
            .Replace("<fonts>", """
                <fonts><extLst><ext uri="urn:test"><font><name val="Extension font" /></font></ext></extLst>
                """)
            .Replace("<cellXfs>", """
                <cellXfs><extLst><ext uri="urn:test"><xf numFmtId="49" /></ext></extLst>
                """);

        var expected = LegacyXlsWriter.StreamingStyles.Create(xml, OfficeTheme.Modern);
        var actual = LegacyXlsWriter.StreamingStyles.Create(decorated, OfficeTheme.Modern);

        Assert.Equal(expected.Fonts.Select(bytes => BitConverter.ToString(bytes)), actual.Fonts.Select(bytes => BitConverter.ToString(bytes)));
        Assert.Equal(expected.Formats.Select(bytes => BitConverter.ToString(bytes)), actual.Formats.Select(bytes => BitConverter.ToString(bytes)));
        Assert.Equal(expected.CellFormats.Select(bytes => BitConverter.ToString(bytes)), actual.CellFormats.Select(bytes => BitConverter.ToString(bytes)));
    }
}
