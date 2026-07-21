using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class OfficeFormatCatalogContractTests {
    [Fact]
    public void CatalogsExposeEveryLegacyAndModernFormatVariant() {
        Assert.Equal(
            new[] { ".doc", ".docm", ".docx", ".dot", ".dotm", ".dotx" },
            WordFormatCatalog.All.Select(format => format.Extension).OrderBy(value => value).ToArray());
        Assert.Equal(
            new[] { ".xla", ".xlam", ".xlm", ".xls", ".xlsb", ".xlsm", ".xlsx", ".xlt", ".xltm", ".xltx", ".xlw" },
            ExcelFormatCatalog.All.Select(format => format.Extension).OrderBy(value => value).ToArray());
        Assert.Equal(
            new[] { ".pot", ".potm", ".potx", ".ppa", ".ppam", ".pps", ".ppsm", ".ppsx", ".ppt", ".pptm", ".pptx" },
            PowerPointFormatCatalog.All.Select(format => format.Extension).OrderBy(value => value).ToArray());
    }

    [Fact]
    public void CatalogsClassifyContainerKindAndMacroCapability() {
        OfficeFormatDescriptor xlsb = ExcelFormatCatalog.GetByExtension("BOOK.XLSB");
        OfficeFormatDescriptor dotm = WordFormatCatalog.GetByExtension(".dotm");
        OfficeFormatDescriptor pps = PowerPointFormatCatalog.GetByExtension("show.pps");

        Assert.Equal(OfficeFormatEncoding.BinaryOpenXml, xlsb.Encoding);
        Assert.True(xlsb.IsMacroEnabled);
        Assert.Equal(OfficeDocumentKind.Template, dotm.DocumentKind);
        Assert.True(dotm.IsMacroEnabled);
        Assert.Equal(OfficeDocumentKind.SlideShow, pps.DocumentKind);
        Assert.Equal(OfficeFormatGeneration.Legacy, pps.Generation);
    }
}
