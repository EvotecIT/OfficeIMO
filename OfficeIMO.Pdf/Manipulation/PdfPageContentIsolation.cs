using System.Threading;

namespace OfficeIMO.Pdf;

/// <summary>Isolates preserved page streams before adding content in the page's default graphics state.</summary>
internal static class PdfPageContentIsolation {
    internal const int WrapperBytes = 5;

    /// <summary>Wraps original stream entries without decoding, repairing, or changing their bytes.</summary>
    internal static PdfArray Isolate(Dictionary<int, PdfIndirectObject> objects, PdfObject contents,
        ref int nextObjectNumber, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        int save = checked(nextObjectNumber++);
        int restore = checked(nextObjectNumber++);
        objects[save] = new PdfIndirectObject(save, 0,
            new PdfStream(new PdfDictionary(), PdfEncoding.Latin1GetBytes("q\n")));
        // This delimiter belongs to the added stream: the original may end at its last operator.
        objects[restore] = new PdfIndirectObject(restore, 0,
            new PdfStream(new PdfDictionary(), PdfEncoding.Latin1GetBytes("\nQ\n")));
        var isolated = new PdfArray();
        isolated.Items.Add(new PdfReference(save, 0));
        if (PdfObjectLookup.Resolve(objects, contents) is PdfArray entries) {
            foreach (PdfObject item in entries.Items) {
                cancellationToken.ThrowIfCancellationRequested();
                isolated.Items.Add(item);
            }
        } else {
            isolated.Items.Add(contents);
        }
        isolated.Items.Add(new PdfReference(restore, 0));
        return isolated;
    }

    /// <summary>Appends a page-space stream after restoring the original page graphics state.</summary>
    /// <returns>The number of added wrapper bytes, or zero when there was no original content.</returns>
    internal static int Append(Dictionary<int, PdfIndirectObject> objects, PdfDictionary page,
        int appendedObjectNumber, ref int nextObjectNumber, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var appended = new PdfReference(appendedObjectNumber, 0);
        if (!page.Items.TryGetValue("Contents", out PdfObject? contents)) {
            page.Items["Contents"] = appended;
            return 0;
        }
        PdfArray isolated = Isolate(objects, contents, ref nextObjectNumber, cancellationToken);
        isolated.Items.Add(appended);
        page.Items["Contents"] = isolated;
        return WrapperBytes;
    }
}
