using System.Security.Cryptography;
using OfficeIMO.IWork;
using OfficeIMO.Word;
using OfficeIMO.Word.IWork;

internal static class NativePagesImageWriterSmoke {
    private const string AcceptedHash = "AA0A26DEBF56C6CDBAABDD9C36A97A3D32331ABEB058BAC0E76D37E43D9E019C";
    private const string EditedHash = "BBEB2B19594862C41CBF4D81B5933F8AEF59FD159D52A52434140CAB88886793";

    internal static void Run() {
        byte[] Image(string name) => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, name));
        var document = IWorkPagesDocument.Create(); document.AddParagraph("Before 😀");
        document.AddImage(Image("created-opaque.png"), 180, 120, "Opaque four-color PNG 😀");
        document.AddParagraph("Between");
        document.AddImage(Image("created-alpha.png"), 120, 80, "Transparent, half-alpha, opaque purple");
        document.AddTable(IWorkTableData.Create(2, 2).SetText(1, 1, "Grid").SetNumber(2, 2, 42), "Values", 100, 28);
        document.AddImage(Image("created-jpeg.jpg"), 160, 80, "Four-color JPEG, stretched frame");
        document.AddParagraph("After");
        byte[] bytes = document.SaveBytes();
        Require(bytes.Length == 15_102 && Convert.ToHexString(SHA256.HashData(bytes)) == AcceptedHash,
            "NativeAOT Pages images differ from the independently decoded Apple-accepted package.");
        Require(bytes.SequenceEqual(document.SaveBytes()), "Native Pages image creation is nondeterministic.");
        VerifySavedWord(bytes, 180);
        using var unchanged = new MemoryStream(); unchanged.WriteByte(9);
        try {
            document.Save(unchanged, new() { MaximumImages = 2 });
            throw new InvalidOperationException("The Pages image budget was not enforced.");
        } catch (InvalidDataException) { }
        Require(unchanged.CanWrite && unchanged.ToArray().SequenceEqual(new byte[] { 9 }),
            "Pages image-budget rejection changed or closed the caller stream.");
        var only = IWorkPagesDocument.Create(); only.AddImage(Image("created-alpha.png"), 120, 80, "Image only");
        using var imageOnly = IWorkSourceDocument.Open(only.SaveBytes()).ToWordDocumentResult(new() { Mode = IWorkConversionMode.EditableOnly });
        imageOnly.RequireCompleteEditableReconstruction();
        Require(imageOnly.Value.Images.Count == 1 && only.Paragraphs.Count == 0, "Image-only Pages body was not editable.");
        Console.WriteLine($"PASS | native Pages inline PNG/JPEG creation | {AcceptedHash} | deterministic bytes, PNG alpha/JPEG, body/table order, point-sized strict DOCX save/reopen, image-only body and image bounds");
        byte[] edited = Image("pages-created-images-edited.pages");
        Require(Convert.ToHexString(SHA256.HashData(edited)) == EditedHash, "Apple-resized Pages image fixture changed.");
        VerifySavedWord(edited, 144);
        Console.WriteLine($"PASS | native Pages resized inline images | {EditedHash} | Apple save/reopen, fractional flow-position cache and point-sized strict DOCX save/reopen");
    }

    private static void VerifySavedWord(byte[] bytes, double firstWidth) {
        var source = IWorkSourceDocument.Open(bytes); var projection = source.ReadPages();
        Require(projection.Images.Count == 3 && projection.Tables.Count == 1 && projection.Diagnostics.Count == 0
            && projection.SourceDeclarationIssues.Count == 0 && projection.SourceReferenceIssues.Count == 0,
            "Pages image projection lost content or native graph declarations.");
        using var result = source.ToWordDocumentResult(new() { Mode = IWorkConversionMode.EditableOnly });
        result.RequireCompleteEditableReconstruction();
        using var output = new MemoryStream(); result.Value.Save(output); output.Position = 0;
        using var saved = WordDocument.Load(output);
        Require(saved.Images.Count == 3 && saved.Tables.Count == 1 && !saved.ValidateDocument().Any(),
            "Pages images or table were lost in saved DOCX.");
        Require(Math.Abs(saved.Images[0].Width!.Value - firstWidth * 96 / 72) < 1e-6
            && Math.Abs(saved.Images[0].Height!.Value - 160) < 1e-6,
            "The inline frame lost its point dimensions in Word's pixel API.");
        for (int index = 0; index < 3; index++) {
            Require(saved.Images[index].ToBytes().SequenceEqual(projection.Images[index].GetBytes())
                && saved.Images[index].Description == projection.Images[index].AccessibilityDescription,
                "Saved DOCX lost an image payload or description.");
        }
    }

    private static void Require(bool condition, string message) {
        if (!condition) throw new InvalidOperationException(message);
    }
}
