using System.Linq;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfDictionaryBoundaryTests {
    [Fact]
    public void Dictionaries_ending_in_a_hex_string_or_holding_angle_brackets_in_strings_read_completely() {
        // Acrobat writes field names as hex strings right before the dictionary's end (<...>>>), and a literal string
        // or a comment may hold >>; none of them ends the dictionary early, directly or inside an object stream.
        string compressed = "7 0 8 45 " +
            "<</FT/Tx/T<4E616D65>/TU(a >> b)/Kids[9 0 R]>>" +
            "<</Fields[7 0 R 10 0 R]/DA(/Helv 0 Tf 0 g)>>";
        int first = compressed.IndexOf('<');
        string pdf = string.Join("\n", new[] {
            "%PDF-1.5",
            "1 0 obj",
            "<</Type/Catalog/Pages 2 0 R/AcroForm 8 0 R>>",
            "endobj",
            "2 0 obj",
            "<</Type/Pages/Count 1/Kids[3 0 R]>>",
            "endobj",
            "3 0 obj",
            "<</Type/Page/Parent 2 0 R/MediaBox[0 0 240 180]/Annots[9 0 R 11 0 R]>>",
            "endobj",
            "6 0 obj",
            $"<</Type/ObjStm/N 2/First {first}/Length {compressed.Length}>>",
            "stream",
            compressed,
            "endstream",
            "endobj",
            "9 0 obj",
            "<</Type/Annot/Subtype/Widget/Parent 7 0 R/Rect[20 100 200 122]/F 4>>",
            "endobj",
            "10 0 obj",
            "<</FT/Btn/T<436865636B>/Kids[11 0 R] % a comment with >> in it",
            ">>",
            "endobj",
            "11 0 obj",
            "<</Type/Annot/Subtype/Widget/Parent 10 0 R/Rect[20 60 34 74]/F 4/AS/Off/AP<</N<</Yes 12 0 R/Off 12 0 R>>>>>>",
            "endobj",
            "12 0 obj",
            "<</Type/XObject/Subtype/Form/BBox[0 0 14 14]/Length 0>>",
            "stream",
            "",
            "endstream",
            "endobj",
            "trailer",
            "<</Root 1 0 R/Size 13>>",
            "%%EOF"
        });

        PdfDocumentInfo info = PdfInspector.Inspect(Encoding.ASCII.GetBytes(pdf));

        Assert.Equal(new[] { "Check", "Name" }, info.FormFields.Select(static field => field.Name).OrderBy(static name => name, StringComparer.Ordinal));
        Assert.Equal("a >> b", info.FormFields.Single(static field => field.Name == "Name").AlternateName);
    }

    [Fact]
    public void A_reference_to_an_undefined_object_reads_as_null_when_the_document_is_rewritten() {
        // ISO 32000-2 7.3.10: an indirect reference to an object that is not defined is the null object, not an error.
        // Producers such as PDFKit leave them behind; encrypting rewrites every object and keeps the reference as null.
        string content = "BT /F1 12 Tf 50 700 Td (Kept) Tj ET";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<</Type/Catalog/Pages 2 0 R/PieceInfo<</Studio<</Data 99 0 R/LastModified(D:20260101000000)>>>>>>",
            "endobj",
            "2 0 obj",
            "<</Type/Pages/Count 1/Kids[3 0 R]>>",
            "endobj",
            "3 0 obj",
            "<</Type/Page/Parent 2 0 R/MediaBox[0 0 300 300]/Resources<</Font<</F1 5 0 R>>>>/Contents 4 0 R>>",
            "endobj",
            "4 0 obj",
            $"<</Length {content.Length}>>",
            "stream",
            content,
            "endstream",
            "endobj",
            "5 0 obj",
            "<</Type/Font/Subtype/Type1/BaseFont/Helvetica>>",
            "endobj",
            "trailer",
            "<</Root 1 0 R/Size 6>>",
            "%%EOF"
        });

        PdfSecurityMutationResult result = PdfDocument.Load(Encoding.ASCII.GetBytes(pdf)).Security
            .Encrypt(new PdfStandardEncryptionOptions("open"));

        PdfDocument encrypted = PdfDocument.Load(result.Pdf, new PdfLoadOptions { Password = "open" });
        Assert.Single(encrypted.Text.Find("Kept"));
        Assert.True(result.PreservationReport.IsPreserved);
    }
}
