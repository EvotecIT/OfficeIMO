using System.IO.Compression;

namespace OfficeIMO.Excel.Xlsb.Write;

internal static partial class XlsbNewPackageWriter {
    internal static void CompleteStreamingWorkbook(ZipArchive archive, IReadOnlyList<string> names,
        ExcelDateSystem dateSystem, byte[] stylesPart, string themeXml) {
        WriteEntry(archive, "[Content_Types].xml", CreateContentTypes(names.Count, true, false, true), CompressionLevel.Fastest);
        WriteEntry(archive, "_rels/.rels", RootRelationships, CompressionLevel.Fastest);
        WriteEntry(archive, "xl/workbook.bin", XlsbWorkbookPartWriter.CreateDirectTabular(names,
            dateSystem == ExcelDateSystem.NineteenFour), CompressionLevel.Fastest);
        WriteEntry(archive, "xl/_rels/workbook.bin.rels", CreateWorkbookRelationships(names.Count, true, false, true), CompressionLevel.Fastest);
        WriteEntry(archive, "xl/styles.bin", stylesPart, CompressionLevel.Fastest);
        WriteEntry(archive, "xl/theme/theme1.xml", themeXml, CompressionLevel.Fastest);
    }
}
