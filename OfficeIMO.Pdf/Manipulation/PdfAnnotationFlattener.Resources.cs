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

}
