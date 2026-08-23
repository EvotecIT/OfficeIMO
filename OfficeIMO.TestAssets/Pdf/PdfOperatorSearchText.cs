using System.Text;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Tests;

/// <summary>
/// Builds assertion text from both PDF syntax and decoded page/form content streams.
/// Rendering tests can therefore inspect operators without coupling production output
/// to an uncompressed serialization choice.
/// </summary>
internal static class PdfOperatorSearchText {
    internal static string From(byte[] pdf) {
        if (pdf == null) throw new ArgumentNullException(nameof(pdf));

        Encoding latin1 = Encoding.GetEncoding(28591);
        var text = new StringBuilder(latin1.GetString(pdf));
        var (objects, _) = PdfSyntax.ParseObjects(pdf);
        PdfReadDocument document = PdfReadDocument.Open(pdf);
        var decodedObjects = new HashSet<int>();

        for (int pageIndex = 0; pageIndex < document.Pages.Count; pageIndex++) {
            int pageObjectNumber = document.Pages[pageIndex].ObjectNumber;
            if (!objects.TryGetValue(pageObjectNumber, out PdfIndirectObject? pageObject) ||
                pageObject.Value is not PdfDictionary pageDictionary ||
                !pageDictionary.Items.TryGetValue("Contents", out PdfObject? contents)) {
                continue;
            }

            AppendStreams(text, objects, contents, decodedObjects, latin1);
        }

        foreach (KeyValuePair<int, PdfIndirectObject> entry in objects) {
            if (entry.Value.Value is not PdfStream stream ||
                !IsFormXObject(stream.Dictionary) ||
                !decodedObjects.Add(entry.Key)) {
                continue;
            }

            AppendDecodedStream(text, objects, stream, latin1);
        }

        return text.ToString();
    }

    internal static string Decode(PdfStream stream, Dictionary<int, PdfIndirectObject> objects) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (objects == null) throw new ArgumentNullException(nameof(objects));

        byte[] decoded = StreamDecoder.Decode(stream.Dictionary, stream.Data, objects);
        return Encoding.GetEncoding(28591).GetString(decoded);
    }

    private static void AppendStreams(
        StringBuilder text,
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject contents,
        HashSet<int> decodedObjects,
        Encoding latin1) {
        if (contents is PdfReference reference) {
            if (decodedObjects.Add(reference.ObjectNumber) &&
                objects.TryGetValue(reference.ObjectNumber, out PdfIndirectObject? indirect) &&
                indirect.Value is PdfStream stream) {
                AppendDecodedStream(text, objects, stream, latin1);
            }

            return;
        }

        if (contents is not PdfArray array) return;

        for (int i = 0; i < array.Items.Count; i++) {
            AppendStreams(text, objects, array.Items[i], decodedObjects, latin1);
        }
    }

    private static void AppendDecodedStream(
        StringBuilder text,
        Dictionary<int, PdfIndirectObject> objects,
        PdfStream stream,
        Encoding latin1) {
        byte[] decoded = StreamDecoder.Decode(stream.Dictionary, stream.Data, objects);
        text.Append('\n').Append(latin1.GetString(decoded));
    }

    private static bool IsFormXObject(PdfDictionary dictionary) =>
        dictionary.Items.TryGetValue("Subtype", out PdfObject? subtype) &&
        subtype is PdfName name &&
        string.Equals(name.Name, "Form", StringComparison.Ordinal);
}
