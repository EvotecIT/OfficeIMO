using System.Globalization;
using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static int AddTrueTypeFontFileObject(IList<byte[]> list, PdfTrueTypeFontProgram fontProgram,
        bool compress, PdfOptions options, string source, CancellationToken cancellationToken) {
        byte[] data = fontProgram.BuildSubsetFontFile(compress, out int uncompressedLength, cancellationToken);
        string filter = compress ? " /Filter /FlateDecode" : string.Empty;
        int id = AddStreamObject(list,
            "<< /Length " + data.Length.ToString(CultureInfo.InvariantCulture) +
            " /Length1 " + uncompressedLength.ToString(CultureInfo.InvariantCulture) + filter + " >>", data);
        if (options.HasDiagnosticsReport) {
            // The cached subset is shared with the writer. Comparing it also identifies the
            // uncommon full-program fallback instead of promising subsetting from byte size alone.
            bool fullProgram = fontProgram.BuildSubsetFontFile().SequenceEqual(fontProgram.FontDataForInspection);
            options.ReportFontEmbedding(source, fontProgram.FontName, "TrueType", id,
                fontProgram.FontDataForInspection.Length, uncompressedLength, data.Length, compress,
                fullProgram ? "full-truetype" : "subset-truetype");
        }
        return id;
    }

    private static int AddOpenTypeCffFontFileObject(IList<byte[]> list, PdfOpenTypeCffFontProgram fontProgram,
        PdfOpenTypeCffCompactFontFile plan, bool compress, PdfOptions options, string source, CancellationToken cancellationToken) {
        byte[] data = compress ? DeflateZlib(plan.Data, cancellationToken) : plan.Data;
        string filter = compress ? " /Filter /FlateDecode" : string.Empty;
        int id = AddStreamObject(list,
            "<< /Length " + data.Length.ToString(CultureInfo.InvariantCulture) +
            " /Subtype /" + plan.FontFileSubtype + filter + " >>", data);
        options.ReportFontEmbedding(source, fontProgram.FontName, "OpenType/CFF", id,
            fontProgram.FontDataLength, plan.Data.Length, data.Length, compress,
            plan.CharStringSubset?.IsSubset == true ? "subset-opentype-cff" :
                plan.IsCompact ? "compact-opentype-cff" : "full-opentype-cff");
        return id;
    }
}
