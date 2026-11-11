using System.IO.Compression;

namespace OfficeIMO.Excel.Xlsb.Write;

internal static partial class XlsbNewPackageWriter {
    internal static void CompleteStreamingWorkbook(ExcelPackageArchive archive, IReadOnlyList<string> names,
        ExcelDateSystem dateSystem, byte[] stylesPart, string themeXml) {
        archive.WriteEntry("[Content_Types].xml", CreateContentTypes(names.Count, true, false, true), CompressionLevel.Fastest);
        archive.WriteEntry("_rels/.rels", RootRelationships, CompressionLevel.Fastest);
        archive.WriteEntry("xl/workbook.bin", XlsbWorkbookPartWriter.CreateDirectTabular(names,
            dateSystem == ExcelDateSystem.NineteenFour), CompressionLevel.Fastest);
        archive.WriteEntry("xl/_rels/workbook.bin.rels", CreateWorkbookRelationships(names.Count, true, false, true), CompressionLevel.Fastest);
        archive.WriteEntry("xl/styles.bin", stylesPart, CompressionLevel.Fastest);
        archive.WriteEntry("xl/theme/theme1.xml", themeXml, CompressionLevel.Fastest);
    }
}
