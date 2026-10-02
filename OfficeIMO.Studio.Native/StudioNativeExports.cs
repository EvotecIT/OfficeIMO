using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Native;

/// <summary>Stateless C ABI for the native Apple prototype. All document operations stay in OfficeIMO.Pdf.</summary>
public static unsafe class StudioNativeExports {
    private const int MaximumInputBytes = 64 * 1024 * 1024;
    private static readonly UTF8Encoding StrictUtf8 = new(false, true);

    /// <summary>
    /// Executes inspect (0), add note (1), or create sample (2). Input is borrowed for this call.
    /// Output is owned by the caller and must be released with oi_studio_free, including error messages.
    /// Returns zero on success, -1 with a UTF-8 error, or -2 when no error buffer can be returned.
    /// </summary>
    [UnmanagedCallersOnly(EntryPoint = "oi_studio_pdf")]
    public static int Execute(int operation, byte* input, int inputLength, int pageNumber,
        byte* note, int noteLength, byte** output, int* outputLength) {
        if (output == null || outputLength == null) return -2;
        *output = null;
        *outputLength = 0;
        try {
            byte[] result;
            if (operation == 2) {
                result = PdfDocument.Create(builder => builder.Page(page => page.Content(content => {
                    content.Text("OFFICEIMO STUDIO", color: new PdfColor(0.2, 0.35, 0.5));
                    content.Spacer(20);
                    content.H1("Review a document");
                    content.Spacer(12);
                    content.Text("Open a PDF, explore its pages, and add a review note.");
                    content.Spacer(24);
                    content.H2("Make space for your notes");
                    content.Text("Choose Add Note to leave a comment on the current page. Read your comments in the page sidebar.");
                    content.Spacer(16);
                    content.Text("Your document stays on your device. Save a Copy exports a separate PDF for your next step.");
                    content.PageBreak();
                    content.H1("A second page");
                    content.Spacer(12);
                    content.Text("Use the page sidebar or the arrows below the document to move between pages.");
                    content.Text("Try adding a note here, then close and reopen the document to read it again.");
                }))).ToBytes();
            } else {
                if (input == null || inputLength <= 0 || inputLength > MaximumInputBytes)
                    throw new ArgumentException("Choose a PDF smaller than 64 MiB.");
                byte[] bytes = new ReadOnlySpan<byte>(input, inputLength).ToArray();
                var document = PdfDocument.Load(bytes);
                if (operation == 0) {
                    int pages = document.Read().Pages.Count;
                    if (pages == 0) throw new ArgumentException("The file does not contain readable PDF pages.");
                    using var stream = new MemoryStream();
                    using (var json = new Utf8JsonWriter(stream)) {
                        json.WriteStartObject(); json.WriteNumber("pages", pages); json.WriteEndObject();
                    }
                    result = stream.ToArray();
                } else if (operation == 1) {
                    if (note == null || noteLength <= 0 || noteLength > 16384)
                        throw new ArgumentException("Enter a note of up to 4,000 characters.");
                    string text = StrictUtf8.GetString(new ReadOnlySpan<byte>(note, noteLength));
                    if (string.IsNullOrWhiteSpace(text) || text.Length > 4000)
                        throw new ArgumentException("Enter a note of up to 4,000 characters.");
                    var pages = document.Read().Pages;
                    if (pageNumber < 1 || pageNumber > pages.Count)
                        throw new ArgumentOutOfRangeException(nameof(pageNumber), "The selected page is unavailable.");
                    var page = pages[pageNumber - 1];
                    var box = page.Geometry.EffectiveBox
                        ?? throw new ArgumentException("The selected page has no usable visible bounds.");
                    double scale = page.UserUnit ?? 1;
                    bool sideways = page.RotationDegrees is 90 or 270;
                    double width = (sideways ? box.Height : box.Width) * scale;
                    double height = (sideways ? box.Width : box.Height) * scale;
                    double margin = Math.Min(36, Math.Min(width, height) / 6);
                    double size = Math.Min(24, Math.Min(width, height) / 3);
                    var rectangle = page.MapVisualRectangleToUserSpace(margin, height - margin - size, margin + size, height - margin);
                    result = document.Annotations.Add(new PdfAnnotationCreateOptions {
                        PageNumber = pageNumber, Subtype = "Text", Contents = text,
                        Title = "OfficeIMO Studio", Rectangle = [rectangle.Left, rectangle.Bottom, rectangle.Right, rectangle.Top],
                        Color = [1, 0.76, 0.18], IconName = "Note", CreatePopup = true
                    }).Bytes;
                } else throw new ArgumentException("Unknown document operation.");
            }
            if (result.Length > MaximumInputBytes) throw new InvalidOperationException("The resulting PDF exceeds the prototype's 64 MiB limit.");
            Write(result, output, outputLength);
            return 0;
        } catch (Exception error) {
            try { Write(Encoding.UTF8.GetBytes(error.Message), output, outputLength); return -1; }
            catch { return -2; }
        }
    }

    private static void Write(byte[] bytes, byte** output, int* length) {
        byte* buffer = (byte*)NativeMemory.Alloc((nuint)bytes.Length);
        bytes.AsSpan().CopyTo(new Span<byte>(buffer, bytes.Length));
        *output = buffer;
        *length = bytes.Length;
    }

    /// <summary>Releases a returned buffer once. Passing null is allowed.</summary>
    [UnmanagedCallersOnly(EntryPoint = "oi_studio_free")]
    public static void Free(void* buffer) => NativeMemory.Free(buffer);
}
