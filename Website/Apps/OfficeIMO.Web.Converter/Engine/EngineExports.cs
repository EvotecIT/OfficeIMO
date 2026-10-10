using System.Diagnostics;
using System.Runtime.InteropServices.JavaScript;
using System.Runtime.Versioning;
using System.Text.Json;
using OfficeIMO.Pdf;
using OfficeIMO.Web.Converter.Services;

namespace OfficeIMO.Web.Converter.Engine;

/// <summary>
/// JavaScript entry points called by engine-worker.js. The worker serializes calls, so the session is single-threaded.
/// Results cross the boundary as JSON; files cross as byte arrays.
/// </summary>
[SupportedOSPlatform("browser")]
public static partial class EngineExports {
    private static readonly ToolSession Session = new();

    public static void Main() { }

    [JSExport]
    internal static void Stage(int slot, byte[] bytes, string fileName) => Session.Stage(slot, bytes, fileName);

    [JSExport]
    internal static void ClearInputs() => Session.ClearInputs();

    /// <summary>Runs one tool action. <paramref name="kind"/> is convert, pdf, origin, or text.</summary>
    [JSExport]
    internal static string Run(string kind, string target, string action, string optionsJson) {
        var stopwatch = Stopwatch.StartNew();
        ToolResultDocument result;
        try {
            ToolOptions options = ToolOptions.Parse(optionsJson);
            result = kind switch {
                "convert" => ConvertTool.Run(Session, target, action, options),
                "pdf" => PdfTool.Run(Session, target, action, options),
                "origin" => OriginTool.Run(Session, action, options),
                "text" => TextTool.Run(Session, action, options),
                _ => throw new NotSupportedException($"Unknown tool kind '{kind}'.")
            };
        } catch (EngineNeedsAssemblyException needs) {
            return JsonSerializer.Serialize(ToolResultDocument.NeedsAssemblies(needs.Assemblies), EngineJsonContext.Default.ToolResultDocument);
        } catch (Exception error) when (error is not OutOfMemoryException) {
            result = ToolResultDocument.Failure(FailureTitle(error), Describe(error));
        }
        stopwatch.Stop();
        if (result.ElapsedMilliseconds == 0) result = result with { ElapsedMilliseconds = stopwatch.ElapsedMilliseconds };
        return JsonSerializer.Serialize(result, EngineJsonContext.Default.ToolResultDocument);
    }

    [JSExport]
    internal static byte[] Artifact(int index) => Session.Artifact(index);

    /// <summary>Reads page count and protection state of a staged PDF without changing it.</summary>
    [JSExport]
    internal static string Probe(int slot) {
        PdfProbeDocument probe;
        try {
            probe = PdfTool.Probe(Session, slot);
        } catch (Exception error) when (error is not OutOfMemoryException) {
            probe = new PdfProbeDocument(false, 0, false, false, false, Describe(error));
        }
        return JsonSerializer.Serialize(probe, EngineJsonContext.Default.PdfProbeDocument);
    }

    /// <summary>Renders one page of a staged input or an artifact as PNG. Returns an empty array when the page can't be drawn.</summary>
    [JSExport]
    internal static byte[] RenderPage(string source, int index, int page, int maximumDimension) {
        try {
            PdfPageRenderResult result = Session.Preview(source, index).Render(page, maximumDimension);
            return result.Succeeded && result.Bytes is { } image ? image : [];
        } catch (Exception error) when (error is not OutOfMemoryException) {
            return [];
        }
    }

    /// <summary>Page count of a staged input or artifact PDF, or 0 when it can't be previewed.</summary>
    [JSExport]
    internal static int PageCount(string source, int index) {
        try {
            return Session.Preview(source, index).PageCount;
        } catch (Exception error) when (error is not OutOfMemoryException) {
            return 0;
        }
    }

    /// <summary>Pays one-time start-up costs (font pack, static tables) while the visitor is still choosing a file.</summary>
    [JSExport]
    internal static void Warmup(string kind) {
        try {
            if (kind is "convert" or "pdf") _ = BrowserPortablePdfProfile.FontPackFingerprint;
        } catch (Exception error) when (error is not OutOfMemoryException) {
            // Warm-up is best effort; the real run reports any failure.
        }
    }

    private static string FailureTitle(Exception error) => error switch {
        _ when LooksDamaged(error) => "The file couldn't be read",
        NotSupportedException => "This file or option isn't supported",
        InvalidOperationException => "Something is missing",
        ArgumentException => "Check the settings",
        InvalidDataException or IOException => "The file couldn't be read",
        _ => "The tool couldn't finish"
    };

    private static string Describe(Exception error) {
        if (error.GetType().Name.Contains("PdfTextEncodingPreflightException", StringComparison.Ordinal)) {
            return "The document uses text that needs a font the browser version doesn't include. " + error.Message;
        }
        // Package readers report damage as a bare code such as "FileContainsCorruptedData".
        return LooksDamaged(error)
            ? "It looks damaged, or it isn't the kind of file its name says. Try opening and re-saving it in the app that made it."
            : error.Message;
    }

    private static bool LooksDamaged(Exception error) =>
        error.Message.Length > 0 && !error.Message.Contains(' ', StringComparison.Ordinal) ||
        error.GetType().Name.Contains("FileFormat", StringComparison.Ordinal) ||
        error.GetType().Name.Contains("OpenXmlPackage", StringComparison.Ordinal);
}
