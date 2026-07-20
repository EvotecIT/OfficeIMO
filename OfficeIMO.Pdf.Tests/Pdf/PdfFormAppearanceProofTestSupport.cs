using System.Text;
using OfficeIMO.Pdf;

namespace OfficeIMO.Tests.Pdf;

internal static class PdfFormAppearanceProofTestSupport {
    public static PdfFormAppearanceProofResult BuildFormAppearanceProof() {
        byte[] source = BuildFormAppearanceProofPdf();
        byte[] filled = PdfFormFiller.FillFields(source, new Dictionary<string, PdfFormFieldValue> {
            ["Name"] = PdfFormFieldValue.From("Visible Value"),
            ["Country"] = PdfFormFieldValue.From("PL"),
            ["AcceptTerms"] = PdfFormFieldValue.From("Yes"),
            ["Payment.Method"] = PdfFormFieldValue.From("Wire"),
            ["Notes"] = PdfFormFieldValue.From("Line one\nLine two"),
            ["Code"] = PdfFormFieldValue.From("ZX91"),
            ["Countries"] = PdfFormFieldValue.FromValues("DE", "US")
        });
        byte[] flattened = PdfFormFiller.FlattenFields(filled);

        return new PdfFormAppearanceProofResult(
            source,
            filled,
            flattened,
            PdfInspector.Inspect(filled),
            PdfInspector.Inspect(flattened),
            Encoding.ASCII.GetString(filled),
            GetFlattenedAppearanceStreamText(flattened));
    }

    public static byte[] BuildFormAppearanceProofPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 320 360] /Contents 4 0 R /Annots [8 0 R 10 0 R 12 0 R 14 0 R 15 0 R 19 0 R 21 0 R 23 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Fields [7 0 R 9 0 R 11 0 R 13 0 R 18 0 R 20 0 R 22 0 R] >>",
            "endobj",
            "7 0 obj",
            "<< /FT /Tx /T (Name) /V (OfficeIMO) /Kids [8 0 R] >>",
            "endobj",
            "8 0 obj",
            "<< /Type /Annot /Subtype /Widget /Parent 7 0 R /Rect [20 140 220 162] /F 4 >>",
            "endobj",
            "9 0 obj",
            "<< /FT /Ch /T (Country) /V (DE) /Opt [[(PL) (Poland)] [(DE) (Germany)] [(US) (United States)]] /Kids [10 0 R] >>",
            "endobj",
            "10 0 obj",
            "<< /Type /Annot /Subtype /Widget /Parent 9 0 R /Rect [20 100 220 122] /F 4 >>",
            "endobj",
            "11 0 obj",
            "<< /FT /Btn /T (AcceptTerms) /V /Off /Kids [12 0 R] >>",
            "endobj",
            "12 0 obj",
            "<< /Type /Annot /Subtype /Widget /Parent 11 0 R /Rect [20 64 36 80] /F 4 /AS /Off >>",
            "endobj",
            "13 0 obj",
            "<< /FT /Btn /T (Payment.Method) /Ff 49152 /V /Card /Kids [14 0 R 15 0 R] >>",
            "endobj",
            "14 0 obj",
            "<< /Type /Annot /Subtype /Widget /Parent 13 0 R /Rect [20 32 36 48] /F 4 /AS /Card /AP << /N << /Off 24 0 R /Card 16 0 R >> >> >>",
            "endobj",
            "15 0 obj",
            "<< /Type /Annot /Subtype /Widget /Parent 13 0 R /Rect [80 32 96 48] /F 4 /AS /Off /AP << /N << /Off 24 0 R /Wire 17 0 R >> >> >>",
            "endobj",
            BuildFormStreamObject(16, "Card selected"),
            BuildFormStreamObject(17, "Wire selected"),
            "18 0 obj",
            "<< /FT /Tx /T (Notes) /V (Initial note) /Ff 4096 /Q 1 /Kids [19 0 R] >>",
            "endobj",
            "19 0 obj",
            "<< /Type /Annot /Subtype /Widget /Parent 18 0 R /Rect [20 270 260 326] /F 4 >>",
            "endobj",
            "20 0 obj",
            "<< /FT /Tx /T (Code) /V (ABCD) /Ff 16777216 /MaxLen 4 /Kids [21 0 R] >>",
            "endobj",
            "21 0 obj",
            "<< /Type /Annot /Subtype /Widget /Parent 20 0 R /Rect [20 236 140 258] /F 4 >>",
            "endobj",
            "22 0 obj",
            "<< /FT /Ch /T (Countries) /V [(DE)] /Ff 2097152 /Opt [[(PL) (Poland)] [(DE) (Germany)] [(US) (United States)]] /Kids [23 0 R] >>",
            "endobj",
            "23 0 obj",
            "<< /Type /Annot /Subtype /Widget /Parent 22 0 R /Rect [20 174 220 224] /F 4 >>",
            "endobj",
            BuildFormStreamObject(24, string.Empty),
            "trailer",
            "<< /Root 1 0 R /Size 25 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static string BuildFormStreamObject(int objectNumber, string text) {
        string content = $"BT /F1 10 Tf 0 0 Td ({text}) Tj ET";
        return string.Join("\n", new[] {
            objectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + " 0 obj",
            $"<< /Type /XObject /Subtype /Form /BBox [0 0 16 16] /Length {content.Length} >>",
            "stream",
            content,
            "endstream",
            "endobj"
        });
    }

    private static string GetFlattenedAppearanceStreamText(byte[] pdf) {
        var (objects, _) = PdfSyntax.ParseObjects(pdf);
        PdfDictionary page = (PdfDictionary)objects.Values.First(indirect =>
            indirect.Value is PdfDictionary dictionary &&
            dictionary.Get<PdfName>("Type")?.Name == "Page").Value;
        PdfDictionary resources = (PdfDictionary)page.Items["Resources"];
        PdfDictionary xObjects = (PdfDictionary)resources.Items["XObject"];
        var builder = new StringBuilder();
        foreach (PdfObject item in xObjects.Items.Values) {
            PdfReference reference = (PdfReference)item;
            PdfStream stream = (PdfStream)objects[reference.ObjectNumber].Value;
            builder.AppendLine(Encoding.ASCII.GetString(stream.Data));
        }

        return builder.ToString();
    }
}

internal sealed class PdfFormAppearanceProofResult {
    public PdfFormAppearanceProofResult(
        byte[] source,
        byte[] filled,
        byte[] flattened,
        PdfDocumentInfo filledInfo,
        PdfDocumentInfo flattenedInfo,
        string filledRaw,
        string flattenedAppearanceText) {
        Source = source;
        Filled = filled;
        Flattened = flattened;
        FilledInfo = filledInfo;
        FlattenedInfo = flattenedInfo;
        FilledRaw = filledRaw;
        FlattenedAppearanceText = flattenedAppearanceText;
    }

    public byte[] Source { get; }

    public byte[] Filled { get; }

    public byte[] Flattened { get; }

    public PdfDocumentInfo FilledInfo { get; }

    public PdfDocumentInfo FlattenedInfo { get; }

    public string FilledRaw { get; }

    public string FlattenedAppearanceText { get; }
}
