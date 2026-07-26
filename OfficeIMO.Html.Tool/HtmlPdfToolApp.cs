using System.Text;
using System.Text.Json;
using OfficeIMO.Drawing;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;

namespace OfficeIMO.Html.Tool;

internal static class HtmlPdfToolApp {
    internal const string Usage = """
OfficeIMO.Html.Tool

Usage:
  officeimo-html convert <input.html|input.mhtml|-> [--input-format html|mhtml] [--output <file|->]
                         [--stylesheet <file.css>] [--base-uri <absolute-uri>]
                         [--font-family <name> --font-regular <file.ttf>]
                         [--font-bold <file.ttf>] [--font-italic <file.ttf>]
                         [--font-bold-italic <file.ttf>]
                         [--max-input-bytes <bytes>] [--max-pages <count>]
                         [--pdf-ua-language <tag>] [--force]
  officeimo-html capabilities [--format text|json]

Local and remote resource reads are disabled by default. Data URIs and bounded MHTML
resources remain available. PDF/UA mode configures and analyzes groundwork; it does not
claim conformance without passing external validator evidence.
""";

    internal static async Task<int> RunAsync(
        string[] args,
        Stream standardInput,
        Stream standardOutput,
        TextWriter standardError,
        CancellationToken cancellationToken = default) {
        try {
            HtmlPdfToolArguments parsed = HtmlPdfToolArguments.Parse(args);
            if (parsed.Command == HtmlPdfToolCommand.Help) {
                await WriteUtf8Async(standardOutput, Usage + Environment.NewLine, cancellationToken).ConfigureAwait(false);
                return 0;
            }
            if (parsed.Command == HtmlPdfToolCommand.Capabilities) {
                await WriteCapabilitiesAsync(standardOutput, parsed.JsonCapabilities, cancellationToken).ConfigureAwait(false);
                return 0;
            }
            return await ConvertAsync(parsed, standardInput, standardOutput, standardError, cancellationToken).ConfigureAwait(false);
        } catch (HtmlPdfToolUsageException exception) {
            await standardError.WriteLineAsync(exception.Message).ConfigureAwait(false);
            await standardError.WriteLineAsync(Usage).ConfigureAwait(false);
            return 2;
        } catch (OperationCanceledException) {
            await standardError.WriteLineAsync("Conversion cancelled.").ConfigureAwait(false);
            return 130;
        } catch (FileNotFoundException exception) {
            await standardError.WriteLineAsync(exception.Message).ConfigureAwait(false);
            return 3;
        } catch (IOException exception) {
            await standardError.WriteLineAsync("I/O failed: " + exception.Message).ConfigureAwait(false);
            return 4;
        } catch (Exception exception) {
            await standardError.WriteLineAsync("Conversion failed: " + exception.GetType().Name + ": " + exception.Message).ConfigureAwait(false);
            return 5;
        }
    }

    private static async Task<int> ConvertAsync(
        HtmlPdfToolArguments arguments,
        Stream standardInput,
        Stream standardOutput,
        TextWriter standardError,
        CancellationToken cancellationToken) {
        byte[] input = await ReadInputAsync(arguments.InputPath!, standardInput, arguments.MaxInputBytes, cancellationToken).ConfigureAwait(false);
        var options = new HtmlPdfSaveOptions {
            MaxPageCount = arguments.MaxPages,
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        };
        if (arguments.BaseUri != null) options.BaseUri = new Uri(arguments.BaseUri, UriKind.Absolute);
        foreach (string stylesheetPath in arguments.StylesheetPaths) {
            byte[] stylesheet = await ReadFileBoundedAsync(stylesheetPath, HtmlPdfToolArguments.MaxStylesheetBytes, cancellationToken).ConfigureAwait(false);
            options.AdditionalStylesheets.Add(Encoding.UTF8.GetString(stylesheet));
        }
        if (arguments.RegularFontPath != null) {
            byte[] regular = await ReadFileBoundedAsync(arguments.RegularFontPath, HtmlPdfToolArguments.MaxFontBytes, cancellationToken).ConfigureAwait(false);
            byte[]? bold = await ReadOptionalFontAsync(arguments.BoldFontPath, cancellationToken).ConfigureAwait(false);
            byte[]? italic = await ReadOptionalFontAsync(arguments.ItalicFontPath, cancellationToken).ConfigureAwait(false);
            byte[]? boldItalic = await ReadOptionalFontAsync(arguments.BoldItalicFontPath, cancellationToken).ConfigureAwait(false);
            ConfigureFontFamily(
                options,
                arguments.FontFamilyName!,
                regular,
                bold,
                italic,
                boldItalic);
        }
        PdfComplianceProfile? complianceProfile = null;
        if (arguments.PdfUaLanguage != null) {
            complianceProfile = PdfComplianceProfile.PdfUa1;
            options.DocumentOptions.UsePdfUa(PdfComplianceProfile.PdfUa1, arguments.PdfUaLanguage);
        }

        PdfDocumentConversionResult conversion;
        using var inputStream = new MemoryStream(input, writable: false);
        if (arguments.ResolveInputFormat() == HtmlPdfToolInputFormat.Mhtml) {
            MhtmlDocument document = await MhtmlDocument.LoadAsync(inputStream, cancellationToken: cancellationToken).ConfigureAwait(false);
            conversion = await document.ToPdfDocumentResultAsync(options, cancellationToken).ConfigureAwait(false);
        } else {
            var documentOptions = new HtmlConversionDocumentOptions {
                BaseUri = options.BaseUri,
                Limits = new HtmlConversionLimits {
                    MaxInputCharacters = (int)Math.Min(arguments.MaxInputBytes, int.MaxValue)
                }
            };
            HtmlConversionDocument document = await HtmlConversionDocument.LoadAsync(
                inputStream,
                documentOptions,
                cancellationToken: cancellationToken).ConfigureAwait(false);
            conversion = await document.ToPdfDocumentResultAsync(options, cancellationToken).ConfigureAwait(false);
        }

        PdfComplianceArtifact? complianceArtifact = complianceProfile.HasValue
            ? conversion.Value.CreateComplianceArtifact(complianceProfile.Value)
            : null;
        await SaveAsync(
            conversion,
            complianceArtifact?.ToBytes(),
            arguments.OutputPath!,
            standardOutput,
            arguments.Force,
            cancellationToken).ConfigureAwait(false);
        foreach (PdfConversionWarning warning in conversion.Report.Warnings) {
            await standardError.WriteLineAsync(
                warning.Severity + " " + warning.Code + " [" + warning.Source + "]: " + warning.Message).ConfigureAwait(false);
        }
        if (complianceProfile.HasValue) {
            PdfComplianceProofReport proof = complianceArtifact!.AssessProof();
            await standardError.WriteLineAsync(
                "PDF/UA readiness: " + proof.ProofStatus + ". " + proof.ExternalProofSummary).ConfigureAwait(false);
            foreach (PdfComplianceRequirement requirement in proof.BlockingRequirements) {
                if (requirement.Id == "pdfua-validation") continue;
                await standardError.WriteLineAsync(
                    "PDF/UA blocker " + requirement.Id + ": " + requirement.Diagnostic).ConfigureAwait(false);
            }
        }
        return conversion.Report.Warnings.Any(warning => warning.Severity == PdfConversionWarningSeverity.Error) ? 6 : 0;
    }

    internal static void ConfigureFontFamily(
        HtmlPdfSaveOptions options,
        string familyName,
        byte[] regular,
        byte[]? bold,
        byte[]? italic,
        byte[]? boldItalic) {
        options.DefaultFontFamily = familyName;
        options.Fonts.Add(familyName, regular, OfficeFontStyle.Regular);
        if (bold != null) options.Fonts.Add(familyName, bold, OfficeFontStyle.Bold);
        if (italic != null) options.Fonts.Add(familyName, italic, OfficeFontStyle.Italic);
        if (boldItalic != null) {
            options.Fonts.Add(
                familyName,
                boldItalic,
                OfficeFontStyle.Bold | OfficeFontStyle.Italic);
        }
        options.FontFamily = new PdfEmbeddedFontFamily(
            familyName,
            regular,
            bold,
            italic,
            boldItalic);
    }

    private static async Task SaveAsync(
        PdfDocumentConversionResult conversion,
        byte[]? exactArtifact,
        string outputPath,
        Stream standardOutput,
        bool force,
        CancellationToken cancellationToken) {
        if (outputPath == "-") {
            if (exactArtifact != null) {
                await standardOutput.WriteAsync(exactArtifact.AsMemory(), cancellationToken).ConfigureAwait(false);
            } else {
                await conversion.SaveAsync(standardOutput, cancellationToken).ConfigureAwait(false);
            }
            return;
        }

        string fullPath = Path.GetFullPath(outputPath);
        if (File.Exists(fullPath) && !force) {
            throw new IOException("Output file '" + fullPath + "' already exists. Use --force to replace it.");
        }
        string? directory = Path.GetDirectoryName(fullPath);
        if (!string.IsNullOrWhiteSpace(directory)) Directory.CreateDirectory(directory);
        string temporaryPath = fullPath + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try {
            if (exactArtifact != null) {
                await File.WriteAllBytesAsync(temporaryPath, exactArtifact, cancellationToken).ConfigureAwait(false);
            } else {
                await conversion.SaveAsync(temporaryPath, cancellationToken).ConfigureAwait(false);
            }
            File.Move(temporaryPath, fullPath, force);
        } finally {
            if (File.Exists(temporaryPath)) File.Delete(temporaryPath);
        }
    }

    private static async Task<byte[]> ReadInputAsync(
        string path,
        Stream standardInput,
        long maximumBytes,
        CancellationToken cancellationToken) {
        if (path == "-") return await ReadBoundedAsync(standardInput, maximumBytes, cancellationToken).ConfigureAwait(false);
        return await ReadFileBoundedAsync(path, maximumBytes, cancellationToken).ConfigureAwait(false);
    }

    private static async Task<byte[]> ReadFileBoundedAsync(string path, long maximumBytes, CancellationToken cancellationToken) {
        string fullPath = Path.GetFullPath(path);
        var info = new FileInfo(fullPath);
        if (!info.Exists) throw new FileNotFoundException("Input file '" + fullPath + "' does not exist.", fullPath);
        if (info.Length > maximumBytes) throw new IOException("Input exceeds the configured byte limit.");
        using var stream = new FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.Read, 81920, true);
        return await ReadBoundedAsync(stream, maximumBytes, cancellationToken).ConfigureAwait(false);
    }

    private static Task<byte[]?> ReadOptionalFontAsync(string? path, CancellationToken cancellationToken) =>
        path == null
            ? Task.FromResult<byte[]?>(null)
            : ReadOptionalFontCoreAsync(path, cancellationToken);

    private static async Task<byte[]?> ReadOptionalFontCoreAsync(string path, CancellationToken cancellationToken) =>
        await ReadFileBoundedAsync(path, HtmlPdfToolArguments.MaxFontBytes, cancellationToken).ConfigureAwait(false);

    private static async Task<byte[]> ReadBoundedAsync(Stream stream, long maximumBytes, CancellationToken cancellationToken) {
        using var buffer = new MemoryStream();
        var chunk = new byte[81920];
        while (true) {
            int read = await stream.ReadAsync(chunk.AsMemory(0, chunk.Length), cancellationToken).ConfigureAwait(false);
            if (read == 0) return buffer.ToArray();
            if (buffer.Length > maximumBytes - read) throw new IOException("Input exceeds the configured byte limit.");
            await buffer.WriteAsync(chunk.AsMemory(0, read), cancellationToken).ConfigureAwait(false);
        }
    }

    private static async Task WriteCapabilitiesAsync(Stream output, bool json, CancellationToken cancellationToken) {
        if (!json) {
            foreach (HtmlRenderCapability capability in HtmlRenderCapabilityCatalog.All) {
                await WriteUtf8Async(
                    output,
                    capability.Id + "\t" + capability.SupportLevel + "\t" + string.Join(",", capability.Features) + Environment.NewLine,
                    cancellationToken).ConfigureAwait(false);
            }
            return;
        }

        using var buffer = new MemoryStream();
        using (var writer = new Utf8JsonWriter(buffer, new JsonWriterOptions { Indented = true })) {
            writer.WriteStartArray();
            foreach (HtmlRenderCapability capability in HtmlRenderCapabilityCatalog.All) {
                writer.WriteStartObject();
                writer.WriteString("id", capability.Id);
                writer.WriteString("area", capability.Area);
                writer.WriteString("kind", capability.Kind.ToString());
                writer.WriteString("supportLevel", capability.SupportLevel.ToString());
                writer.WriteString("behavior", capability.Behavior);
                writer.WriteStartArray("features");
                foreach (string feature in capability.Features) writer.WriteStringValue(feature);
                writer.WriteEndArray();
                writer.WriteStartArray("diagnosticCodes");
                foreach (string code in capability.DiagnosticCodes) writer.WriteStringValue(code);
                writer.WriteEndArray();
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
        }
        await output.WriteAsync(buffer.ToArray().AsMemory(), cancellationToken).ConfigureAwait(false);
    }

    private static async Task WriteUtf8Async(Stream output, string text, CancellationToken cancellationToken) {
        byte[] bytes = Encoding.UTF8.GetBytes(text);
        await output.WriteAsync(bytes.AsMemory(), cancellationToken).ConfigureAwait(false);
    }
}
