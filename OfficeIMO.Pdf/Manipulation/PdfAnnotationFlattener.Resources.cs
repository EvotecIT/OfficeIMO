namespace OfficeIMO.Pdf;

internal static partial class PdfAnnotationFlattener {
    private static PdfDictionary EnsurePageXObjects(Dictionary<int, PdfIndirectObject> objects, PdfDictionary page) {
        return PdfPageResourceHelper.EnsurePageXObjects(objects, page, "visual annotation flattening");
    }

    private static string CreateUniqueXObjectName(PdfDictionary xObjects) {
        int index = 1;
        string name;
        do {
            name = "OfficeIMOAnnot" + index.ToString(System.Globalization.CultureInfo.InvariantCulture);
            index++;
        } while (xObjects.Items.ContainsKey(name));

        return name;
    }

    private static PdfStream CreateContentStream(string content) {
        var dictionary = new PdfDictionary();
        return new PdfStream(dictionary, PdfEncoding.Latin1GetBytes(content));
    }

    private static void AppendPageContent(Dictionary<int, PdfIndirectObject> objects, PdfDictionary page, int contentObjectNumber) {
        var newReference = new PdfReference(contentObjectNumber, 0);
        if (!page.Items.TryGetValue("Contents", out var contents)) {
            page.Items["Contents"] = newReference;
            return;
        }

        if (contents is PdfArray contentsArray) {
            contentsArray.Items.Add(newReference);
            return;
        }

        var array = new PdfArray();
        AppendContentEntries(objects, array, contents);
        array.Items.Add(newReference);
        page.Items["Contents"] = array;
    }

    private static void AppendContentEntries(Dictionary<int, PdfIndirectObject> objects, PdfArray target, PdfObject contents) {
        if (contents is PdfArray directArray) {
            foreach (var item in directArray.Items) {
                target.Items.Add(item);
            }

            return;
        }

        if (contents is PdfReference reference &&
            PdfObjectLookup.TryGet(objects, reference, out var indirect) &&
            indirect.Value is PdfArray referencedArray) {
            foreach (var item in referencedArray.Items) {
                target.Items.Add(item);
            }

            return;
        }

        target.Items.Add(contents);
    }
}
