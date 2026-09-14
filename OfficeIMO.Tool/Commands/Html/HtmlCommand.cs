using System.Text;
using System.Text.Json;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Mhtml;
using OfficeIMO.Pdf;

namespace OfficeIMO.Tool.Commands.Html;

internal static class HtmlCommand {
    internal const string Usage = """
OfficeIMO.Tool - HTML

Usage:
  officeimo html convert <input.html|input.mhtml|-> [--input-format html|mhtml] [--output <file|->]
                         [--stylesheet <file.css>] [--base-uri <absolute-uri>]
                         [--font-family <name> --font-regular <file.ttf>]
                         [--font-bold <file.ttf>] [--font-italic <file.ttf>]
                         [--font-bold-italic <file.ttf>]
                         [--max-input-bytes <bytes>] [--max-pages <count>]
                         [--pdf-ua-language <tag>] [--force]
  officeimo html capabilities [--format text|json]

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
            HtmlArguments parsed = HtmlArguments.Parse(args);
            if (parsed.Command == HtmlCommandKind.Help) {
                await WriteUtf8Async(standardOutput, Usage + Environment.NewLine, cancellationToken).ConfigureAwait(false);
                return (int)OfficeImoToolExitCode.Success;
            }
            if (parsed.Command == HtmlCommandKind.Capabilities) {
                await WriteCapabilitiesAsync(standardOutput, parsed.JsonCapabilities, cancellationToken).ConfigureAwait(false);
                return (int)OfficeImoToolExitCode.Success;
            }
            return await ConvertAsync(parsed, standardInput, standardOutput, standardError, cancellationToken).ConfigureAwait(false);
        } catch (HtmlUsageException exception) {
            await standardError.WriteLineAsync(exception.Message).ConfigureAwait(false);
            await standardError.WriteLineAsync(Usage).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Usage;
        } catch (OperationCanceledException) {
            await standardError.WriteLineAsync("Conversion cancelled.").ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Cancelled;
        } catch (FileNotFoundException exception) {
            await standardError.WriteLineAsync(exception.Message).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.InputNotFound;
        } catch (IOException exception) {
            await standardError.WriteLineAsync("I/O failed: " + exception.Message).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.UnsupportedInput;
        } catch (Exception exception) {
            await standardError.WriteLineAsync("Conversion failed: " + exception.GetType().Name + ": " + exception.Message).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.OperationFailed;
        }
    }

    private static async Task<int> ConvertAsync(
        HtmlArguments arguments,
        Stream standardInput,
        Stream standardOutput,
        TextWriter standardError,
        CancellationToken cancellationToken) {
        byte[] input = await ReadInputAsync(arguments.InputPath!, standardInput, arguments.MaxInputBytes, cancellationToken).ConfigureAwait(false);
        var options = new HtmlToPdfOptions {
            MaxPageCount = arguments.MaxPages,
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        };
        if (arguments.BaseUri != null) options.BaseUri = new Uri(arguments.BaseUri, UriKind.Absolute);
        foreach (string stylesheetPath in arguments.StylesheetPaths) {
            byte[] stylesheet = await ReadFileBoundedAsync(stylesheetPath, HtmlArguments.MaxStylesheetBytes, cancellationToken).ConfigureAwait(false);
            options.AdditionalStylesheets.Add(Encoding.UTF8.GetString(stylesheet));
        }
        if (arguments.RegularFontPath != null) {
            byte[] regular = await ReadFileBoundedAsync(arguments.RegularFontPath, HtmlArguments.MaxFontBytes, cancellationToken).ConfigureAwait(false);
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
            options.PdfOptions.UsePdfUa(PdfComplianceProfile.PdfUa1, arguments.PdfUaLanguage);
        }

        PdfDocumentConversionResult conversion;
        using var inputStream = new MemoryStream(input, writable: false);
        if (arguments.ResolveInputFormat() == HtmlInputFormat.Mhtml) {
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
        try {
            await SaveAsync(
                conversion,
                complianceArtifact?.ToBytes(),
                arguments.OutputPath!,
                standardOutput,
                arguments.Force,
                cancellationToken).ConfigureAwait(false);
        } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) {
            await standardError.WriteLineAsync("Output failed: " + exception.Message).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.OutputFailed;
        }
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
        return conversion.Report.Warnings.Any(warning => warning.Severity == PdfConversionWarningSeverity.Error)
            ? (int)OfficeImoToolExitCode.OutputFailed
            : (int)OfficeImoToolExitCode.Success;
    }

    internal static void ConfigureFontFamily(
        HtmlToPdfOptions options,
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
        await ReadFileBoundedAsync(path, HtmlArguments.MaxFontBytes, cancellationToken).ConfigureAwait(false);

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
        IReadOnlyList<string> validationErrors = HtmlRenderCapabilityCatalog.Validate();
        IReadOnlyList<string> renderProfileErrors = HtmlRenderProfileContracts.Validate();
        if (validationErrors.Count != 0 || renderProfileErrors.Count != 0) {
            throw new InvalidOperationException("The HTML capability catalog is invalid: "
                + string.Join(" ", validationErrors.Concat(renderProfileErrors)));
        }

        if (!json) {
            await WriteUtf8Async(output, "schemaVersion\t" + HtmlRenderCapabilityCatalog.SchemaVersion + Environment.NewLine, cancellationToken).ConfigureAwait(false);
            foreach (HtmlCapabilityProfileManifest profile in HtmlRenderCapabilityCatalog.ProfileManifests) {
                await WriteUtf8Async(
                    output,
                    "profile\t" + profile.Id + "\t" + profile.Version + "\t" + profile.Promotion + Environment.NewLine,
                    cancellationToken).ConfigureAwait(false);
            }
            foreach (HtmlRenderProfileContract profile in HtmlRenderProfileContracts.All) {
                await WriteUtf8Async(
                    output,
                    "renderProfile\t" + profile.Id + "\t" + profile.CssMedia + "\t" + profile.Surface + "\t"
                        + profile.Pagination + "\t" + profile.DefaultPageSet.Mode + "\t" + profile.Coverage + "\t"
                        + profile.Promotion + "\t" + string.Join(",", profile.Encoders) + Environment.NewLine,
                    cancellationToken).ConfigureAwait(false);
            }
            foreach (HtmlRenderCapability capability in HtmlRenderCapabilityCatalog.All) {
                foreach (HtmlCapabilityProfileBinding binding in capability.ProfileBindings) {
                    await WriteUtf8Async(
                        output,
                        "capability\t" + capability.Id + "\t" + capability.Kind + "\t" + capability.Stages + "\t"
                            + binding.ProfileId + "\t" + binding.Coverage + "\t" + binding.Handling + "\t"
                            + binding.Maturity + "\t" + binding.Promotion + "\t" + string.Join(",", binding.ProviderIds)
                            + "\t" + string.Join(",", binding.OptionalProviderIds) + Environment.NewLine,
                        cancellationToken).ConfigureAwait(false);
                }
            }
            return;
        }

        using var buffer = new MemoryStream();
        using (var writer = new Utf8JsonWriter(buffer, new JsonWriterOptions { Indented = true })) {
            writer.WriteStartObject();
            writer.WriteNumber("schemaVersion", HtmlRenderCapabilityCatalog.SchemaVersion);
            writer.WriteStartArray("profiles");
            foreach (HtmlCapabilityProfileManifest profile in HtmlRenderCapabilityCatalog.ProfileManifests) {
                writer.WriteStartObject();
                writer.WriteString("id", profile.Id);
                writer.WriteString("version", profile.Version);
                writer.WriteString("title", profile.Title);
                writer.WriteString("promotion", profile.Promotion.ToString());
                WriteStringArray(writer, "platforms", profile.Platforms);
                WriteStringArray(writer, "outputs", profile.Outputs);
                writer.WriteStartArray("providers");
                foreach (HtmlCapabilityProviderPin provider in profile.Providers) {
                    writer.WriteStartObject();
                    writer.WriteString("id", provider.Id);
                    writer.WriteString("name", provider.Name);
                    writer.WriteString("version", provider.Version);
                    writer.WriteString("ownership", provider.Ownership.ToString());
                    writer.WriteEndObject();
                }
                writer.WriteEndArray();
                writer.WriteStartArray("specifications");
                foreach (HtmlCapabilitySpecificationPin specification in profile.Specifications) {
                    writer.WriteStartObject();
                    writer.WriteString("id", specification.Id);
                    writer.WriteString("title", specification.Title);
                    writer.WriteString("uri", specification.Uri);
                    writer.WriteString("revision", specification.Revision);
                    writer.WriteString("scope", specification.Scope);
                    writer.WriteEndObject();
                }
                writer.WriteEndArray();
                writer.WriteStartArray("evidence");
                foreach (HtmlCapabilityEvidencePin evidence in profile.Evidence) {
                    writer.WriteStartObject();
                    writer.WriteString("id", evidence.Id);
                    writer.WriteString("source", evidence.Source);
                    writer.WriteString("revision", evidence.Revision);
                    writer.WriteString("role", evidence.Role.ToString());
                    writer.WriteString("scope", evidence.Scope);
                    WriteNullableNumber(writer, "required", evidence.Required);
                    WriteNullableNumber(writer, "passed", evidence.Passed);
                    WriteNullableNumber(writer, "failed", evidence.Failed);
                    WriteNullableNumber(writer, "excluded", evidence.Excluded);
                    WriteNullableNumber(writer, "untested", evidence.Untested);
                    WriteStringArray(writer, "caseIds", evidence.CaseIds);
                    writer.WriteEndObject();
                }
                writer.WriteEndArray();
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
            writer.WriteStartArray("renderProfiles");
            foreach (HtmlRenderProfileContract profile in HtmlRenderProfileContracts.All) {
                writer.WriteStartObject();
                writer.WriteString("id", profile.Id);
                writer.WriteString("name", profile.Name);
                WriteStringArray(writer, "documentStates", Enum.GetValues(typeof(HtmlRenderDocumentState))
                    .Cast<HtmlRenderDocumentState>().Select(value => value.ToString()));
                writer.WriteString("cssMedia", profile.CssMedia.ToString());
                writer.WriteString("surface", profile.Surface.ToString());
                writer.WriteString("pagination", profile.Pagination.ToString());
                writer.WriteString("defaultPageSet", profile.DefaultPageSet.Mode.ToString());
                writer.WriteString("coverage", profile.Coverage.ToString());
                writer.WriteString("promotion", profile.Promotion.ToString());
                WriteStringArray(writer, "encoders", profile.Encoders.Select(value => value.ToString()));
                WriteStringArray(writer, "pageSets", profile.PageSets.Select(value => value.ToString()));
                WriteStringArray(writer, "capabilityProfileIds", profile.CapabilityProfileIds);
                WriteStringArray(writer, "evidenceIds", profile.EvidenceIds);
                writer.WriteString("behavior", profile.Behavior);
                writer.WriteString("limitations", profile.Limitations);
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
            writer.WriteStartArray("capabilities");
            foreach (HtmlRenderCapability capability in HtmlRenderCapabilityCatalog.All) {
                writer.WriteStartObject();
                writer.WriteString("id", capability.Id);
                writer.WriteString("area", capability.Area);
                writer.WriteString("kind", capability.Kind.ToString());
                WriteStringArray(writer, "stages", Enum.GetValues(typeof(HtmlCapabilityStage)).Cast<HtmlCapabilityStage>()
                    .Where(stage => stage != HtmlCapabilityStage.None && capability.Stages.HasFlag(stage))
                    .Select(stage => stage.ToString()));
                writer.WriteString("behavior", capability.Behavior);
                WriteStringArray(writer, "features", capability.Features);
                WriteStringArray(writer, "limitations", capability.Limitations);
                WriteStringArray(writer, "diagnosticCodes", capability.DiagnosticCodes);
                writer.WriteStartArray("profileBindings");
                foreach (HtmlCapabilityProfileBinding binding in capability.ProfileBindings) {
                    writer.WriteStartObject();
                    writer.WriteString("profileId", binding.ProfileId);
                    writer.WriteString("coverage", binding.Coverage.ToString());
                    writer.WriteString("handling", binding.Handling.ToString());
                    writer.WriteString("maturity", binding.Maturity.ToString());
                    writer.WriteString("promotion", binding.Promotion.ToString());
                    WriteStringArray(writer, "providerIds", binding.ProviderIds);
                    WriteStringArray(writer, "optionalProviderIds", binding.OptionalProviderIds);
                    WriteStringArray(writer, "specificationIds", binding.SpecificationIds);
                    WriteStringArray(writer, "evidenceIds", binding.EvidenceIds);
                    writer.WriteEndObject();
                }
                writer.WriteEndArray();
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
            writer.WriteEndObject();
        }
        await output.WriteAsync(buffer.ToArray().AsMemory(), cancellationToken).ConfigureAwait(false);
    }

    private static void WriteStringArray(Utf8JsonWriter writer, string propertyName, IEnumerable<string> values) {
        writer.WriteStartArray(propertyName);
        foreach (string value in values) writer.WriteStringValue(value);
        writer.WriteEndArray();
    }

    private static void WriteNullableNumber(Utf8JsonWriter writer, string propertyName, int? value) {
        if (value.HasValue) writer.WriteNumber(propertyName, value.Value);
        else writer.WriteNull(propertyName);
    }

    private static async Task WriteUtf8Async(Stream output, string text, CancellationToken cancellationToken) {
        byte[] bytes = Encoding.UTF8.GetBytes(text);
        await output.WriteAsync(bytes.AsMemory(), cancellationToken).ConfigureAwait(false);
    }
}
