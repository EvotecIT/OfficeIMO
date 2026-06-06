using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Text;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfReaderAndFooterRegressionTests {

    private static byte[] BuildPdfWithFormResourceNameEscapes(bool dictionaryUsesEscapedName, bool contentUsesEscapedName) {
        const string formContent = "BT\n/F1 12 Tf\n10 20 Td\n(Escaped form) Tj\nET\n";
        string pageContentName = contentUsesEscapedName ? "/Fm#31" : "/Fm1";
        string resourceName = dictionaryUsesEscapedName ? "/Fm#31" : "/Fm1";
        string pageContent = $"q\n{pageContentName} Do\nQ\n";
        int pageStreamLength = Encoding.ASCII.GetByteCount(pageContent);
        int formStreamLength = Encoding.ASCII.GetByteCount(formContent);

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>",
            "endobj",
            "3 0 obj",
            $"<< /Type /Page /Parent 2 0 R /Resources << /XObject << {resourceName} 5 0 R >> >> /Contents 6 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            "endobj",
            "5 0 obj",
            $"<< /Type /XObject /Subtype /Form /BBox [0 0 100 100] /Resources << /Font << /F1 4 0 R >> >> /Length {formStreamLength} >>",
            "stream",
            formContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "6 0 obj",
            $"<< /Length {pageStreamLength} >>",
            "stream",
            pageContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }


    private static byte[] BuildPdfWithInlineNestedFormResources() {
        const string pageContent = "q\n/Fx Do\nQ\n";
        const string formContent = "BT\n/F1 12 Tf\n10 20 Td\n(Inline form) Tj\nET\n";
        int pageStreamLength = Encoding.ASCII.GetByteCount(pageContent);
        int formStreamLength = Encoding.ASCII.GetByteCount(formContent);

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /Resources << /XObject << /Fx 5 0 R >> >> /Contents 6 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            "endobj",
            "5 0 obj",
            $"<< /Type /XObject /Subtype /Form /BBox [0 0 100 100] /Resources << /Font << /F1 4 0 R >> >> /Length {formStreamLength} >>",
            "stream",
            formContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "6 0 obj",
            $"<< /Length {pageStreamLength} >>",
            "stream",
            pageContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }


    private static byte[] BuildPdfWithInheritedFormResources() {
        const string pageContent = "q\n/Fx Do\nQ\n";
        const string formContent = "BT\n/F1 12 Tf\n10 20 Td\n(Form hello) Tj\nET\n";
        int pageStreamLength = Encoding.ASCII.GetByteCount(pageContent);
        int formStreamLength = Encoding.ASCII.GetByteCount(formContent);

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources 7 0 R >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /Contents 6 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            "endobj",
            "5 0 obj",
            $"<< /Type /XObject /Subtype /Form /BBox [0 0 100 100] /Matrix [1 0 0 1 100 200] /Resources 9 0 R /Length {formStreamLength} >>",
            "stream",
            formContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "6 0 obj",
            $"<< /Length {pageStreamLength} >>",
            "stream",
            pageContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "7 0 obj",
            "<< /XObject 8 0 R >>",
            "endobj",
            "8 0 obj",
            "<< /Fx 5 0 R >>",
            "endobj",
            "9 0 obj",
            "<< /Font 10 0 R >>",
            "endobj",
            "10 0 obj",
            "<< /F1 4 0 R >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildPdfWithRepeatedFormInvocations() {
        const string pageContent = "q\n1 0 0 1 0 0 cm\n/Fx Do\nQ\nq\n1 0 0 1 100 0 cm\n/Fx Do\nQ\n";
        const string formContent = "BT\n/F1 12 Tf\n10 20 Td\n(Repeated form) Tj\nET\n";
        int pageStreamLength = Encoding.ASCII.GetByteCount(pageContent);
        int formStreamLength = Encoding.ASCII.GetByteCount(formContent);

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources 7 0 R >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /Contents 6 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            "endobj",
            "5 0 obj",
            $"<< /Type /XObject /Subtype /Form /BBox [0 0 100 100] /Resources 9 0 R /Length {formStreamLength} >>",
            "stream",
            formContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "6 0 obj",
            $"<< /Length {pageStreamLength} >>",
            "stream",
            pageContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "7 0 obj",
            "<< /XObject 8 0 R >>",
            "endobj",
            "8 0 obj",
            "<< /Fx 5 0 R >>",
            "endobj",
            "9 0 obj",
            "<< /Font 10 0 R >>",
            "endobj",
            "10 0 obj",
            "<< /F1 4 0 R >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildPdfWithNestedFormInvocations() {
        const string pageContent = "q\n1 0 0 1 100 200 cm\n/FxOuter Do\nQ\n";
        const string outerFormContent = "q\n1 0 0 1 15 25 cm\n/FxInner Do\nQ\n";
        const string innerFormContent = "BT\n/F1 12 Tf\n5 7 Td\n(Nested form) Tj\nET\n";
        int pageStreamLength = Encoding.ASCII.GetByteCount(pageContent);
        int outerFormLength = Encoding.ASCII.GetByteCount(outerFormContent);
        int innerFormLength = Encoding.ASCII.GetByteCount(innerFormContent);

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources 7 0 R >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /Contents 8 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            "endobj",
            "5 0 obj",
            $"<< /Type /XObject /Subtype /Form /BBox [0 0 200 200] /Resources 9 0 R /Length {outerFormLength} >>",
            "stream",
            outerFormContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "6 0 obj",
            $"<< /Type /XObject /Subtype /Form /BBox [0 0 50 50] /Resources 11 0 R /Length {innerFormLength} >>",
            "stream",
            innerFormContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "7 0 obj",
            "<< /XObject 10 0 R >>",
            "endobj",
            "8 0 obj",
            $"<< /Length {pageStreamLength} >>",
            "stream",
            pageContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "9 0 obj",
            "<< /XObject 12 0 R >>",
            "endobj",
            "10 0 obj",
            "<< /FxOuter 5 0 R >>",
            "endobj",
            "11 0 obj",
            "<< /Font 13 0 R >>",
            "endobj",
            "12 0 obj",
            "<< /FxInner 6 0 R >>",
            "endobj",
            "13 0 obj",
            "<< /F1 4 0 R >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildPdfWithFormXObjectImage() {
        const string pageContent = "q\n1 0 0 1 100 200 cm\n/Fx Do\nQ\n";
        const string formContent = "q\n10 0 0 10 0 0 cm\n/ImNested Do\nQ\n";
        const string imageBytes = "abc";
        int pageStreamLength = Encoding.ASCII.GetByteCount(pageContent);
        int formStreamLength = Encoding.ASCII.GetByteCount(formContent);
        int imageLength = Encoding.ASCII.GetByteCount(imageBytes);

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources 4 0 R >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /Contents 5 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /XObject << /Fx 6 0 R >> >>",
            "endobj",
            "5 0 obj",
            $"<< /Length {pageStreamLength} >>",
            "stream",
            pageContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "6 0 obj",
            $"<< /Type /XObject /Subtype /Form /BBox [0 0 20 20] /Resources << /XObject << /ImNested 7 0 R >> >> /Length {formStreamLength} >>",
            "stream",
            formContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "7 0 obj",
            $"<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length {imageLength} >>",
            "stream",
            imageBytes,
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildPdfWithUnusedFormXObjectImage() {
        const string pageContent = "q\nQ\n";
        const string formContent = "q\n10 0 0 10 0 0 cm\n/ImUnused Do\nQ\n";
        const string imageBytes = "abc";
        int pageStreamLength = Encoding.ASCII.GetByteCount(pageContent);
        int formStreamLength = Encoding.ASCII.GetByteCount(formContent);
        int imageLength = Encoding.ASCII.GetByteCount(imageBytes);

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources 4 0 R >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /Contents 5 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /XObject << /FxUnused 6 0 R >> >>",
            "endobj",
            "5 0 obj",
            $"<< /Length {pageStreamLength} >>",
            "stream",
            pageContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "6 0 obj",
            $"<< /Type /XObject /Subtype /Form /BBox [0 0 20 20] /Resources << /XObject << /ImUnused 7 0 R >> >> /Length {formStreamLength} >>",
            "stream",
            formContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "7 0 obj",
            $"<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length {imageLength} >>",
            "stream",
            imageBytes,
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildPdfWithUnusedPageImageResource() {
        const string pageContent = "q\nQ\n";
        const string imageBytes = "abc";
        int pageStreamLength = Encoding.ASCII.GetByteCount(pageContent);
        int imageLength = Encoding.ASCII.GetByteCount(imageBytes);

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources 4 0 R >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /Contents 5 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /XObject << /ImUnused 6 0 R >> >>",
            "endobj",
            "5 0 obj",
            $"<< /Length {pageStreamLength} >>",
            "stream",
            pageContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "6 0 obj",
            $"<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length {imageLength} >>",
            "stream",
            imageBytes,
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildPdfWithScaledFormMatrix() {
        const string pageContent = "q\n2 0 0 2 10 20 cm\n/Fx Do\nQ\n";
        const string formContent = "BT\n/F1 12 Tf\n3 4 Td\n(Scaled form) Tj\nET\n";
        int pageStreamLength = Encoding.ASCII.GetByteCount(pageContent);
        int formStreamLength = Encoding.ASCII.GetByteCount(formContent);

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources 7 0 R >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /Contents 6 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            "endobj",
            "5 0 obj",
            $"<< /Type /XObject /Subtype /Form /BBox [0 0 100 100] /Matrix [1 0 0 1 5 7] /Resources 9 0 R /Length {formStreamLength} >>",
            "stream",
            formContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "6 0 obj",
            $"<< /Length {pageStreamLength} >>",
            "stream",
            pageContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "7 0 obj",
            "<< /XObject 8 0 R >>",
            "endobj",
            "8 0 obj",
            "<< /Fx 5 0 R >>",
            "endobj",
            "9 0 obj",
            "<< /Font 10 0 R >>",
            "endobj",
            "10 0 obj",
            "<< /F1 4 0 R >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildPdfWithInlineFormTextOrdering() {
        const string pageContent = "BT /F1 12 Tf (Before ) Tj /Fx Do ( after) Tj ET\n";
        const string formContent = "BT /F1 12 Tf (middle) Tj ET\n";
        int pageStreamLength = Encoding.ASCII.GetByteCount(pageContent);
        int formStreamLength = Encoding.ASCII.GetByteCount(formContent);

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources 7 0 R >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /Contents 5 0 R >>",
            "endobj",
            "4 0 obj",
            $"<< /Type /XObject /Subtype /Form /BBox [0 0 100 100] /Resources 8 0 R /Length {formStreamLength} >>",
            "stream",
            formContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "5 0 obj",
            $"<< /Length {pageStreamLength} >>",
            "stream",
            pageContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "7 0 obj",
            "<< /Font 9 0 R /XObject 10 0 R >>",
            "endobj",
            "8 0 obj",
            "<< /Font 9 0 R >>",
            "endobj",
            "9 0 obj",
            "<< /F1 6 0 R >>",
            "endobj",
            "10 0 obj",
            "<< /Fx 4 0 R >>",
            "endobj",
            "6 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static PdfReadPage CreatePdfReadPageWithDirectFormCycle() {
        var state = BuildDirectFormCycleState();
        var pageDict = new PdfDictionary();
        pageDict.Items["Type"] = new PdfName("Page");
        pageDict.Items["Resources"] = state.Resources;

        var contents = new PdfArray();
        contents.Items.Add(new PdfStream(new PdfDictionary(), Encoding.ASCII.GetBytes("/Fx Do")));
        pageDict.Items["Contents"] = contents;

        var ctor = typeof(PdfReadPage).GetConstructor(
            BindingFlags.Instance | BindingFlags.NonPublic,
            binder: null,
            new[] { typeof(int), typeof(PdfDictionary), typeof(Dictionary<int, PdfIndirectObject>) },
            modifiers: null);

        return Assert.IsType<PdfReadPage>(ctor!.Invoke(new object[] { 1, pageDict, state.Objects }));
    }

    private static (PdfDictionary Resources, Dictionary<int, PdfIndirectObject> Objects) BuildDirectFormCycleState() {
        var xObjects = new PdfDictionary();
        var resources = new PdfDictionary();
        resources.Items["XObject"] = xObjects;

        var formDict = new PdfDictionary();
        formDict.Items["Subtype"] = new PdfName("Form");
        formDict.Items["Resources"] = resources;

        var formStream = new PdfStream(formDict, Encoding.ASCII.GetBytes("/Fx Do"));
        xObjects.Items["Fx"] = formStream;

        return (resources, new Dictionary<int, PdfIndirectObject>());
    }

}
