using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Serialization;
using OfficeIMO.Pdf;
using OfficeIMO.Internal;

namespace OfficeIMO.Workflows;

internal static partial class OfficeConversionBatchExecutor {
    private static readonly JsonSerializerOptions ConfigurationJson = new() { IncludeFields = true,
        Converters = { new PdfConfigurationConverter() } };
    private static readonly OfficeConversionConfigurationJsonContext ConfigurationContext = new(ConfigurationJson);

    private static string CaptureConfiguration(OfficeConversionBatchRequest settings) => Hash(string.Join("\n", Schema,
        settings.InputDirectory == null ? "selected-files" : OfficePathIdentity.GetPathIdentityKey(settings.InputDirectory),
        OfficePathIdentity.GetPathIdentityKey(settings.OutputDirectory), settings.TargetExtension));

    private static string CaptureRenderingConfiguration(OfficeConversionBatchRequest settings, string routeId) {
        OfficeWorkflowConversionOptions options = settings.ConversionOptions.ForRoute(routeId);
        if (options.Html?.ResourceResolver != null || options.Html?.TextShapingProvider != null || options.Rtf?.ImageConverter != null || options.Markdown?.RemoteImageResolver != null ||
            options.PublisherRead?.ImageCodec != null ||
            options.Visio?.DrawingOptions?.TextShapingProvider != null ||
            options.Visio?.DrawingOptions?.Fonts.FontProgramProvider != null || options.Visio?.DrawingOptions?.Fonts.FontVariationResolver != null ||
            options.Markdown?.ResourcePolicy.AllowRemoteResourceResolution == true ||
            options.Html?.ResourcePolicy.AllowRemoteResourceResolution == true)
            throw new NotSupportedException("Runtime resource callbacks and remote resources require an ordinary batch without checkpoints.");
        if (options.Markdown?.ResourcePolicy.AllowLocalFileAccess == true && !options.Markdown.RestrictLocalImagesToBaseDirectory)
            throw new NotSupportedException("Checkpointed Markdown resources must remain inside BaseDirectory. Unrestricted local resources require an ordinary batch without checkpoints.");
        using var hash = SHA256.Create();
        using (var stream = new CryptoStream(Stream.Null, hash, CryptoStreamMode.Write)) {
            JsonSerializer.Serialize(stream, new OfficeConversionRenderingConfiguration(routeId, settings.OutputProfile, options,
                Environment.MachineName, routeId is "odg-pdf" or "visio-pdf" or "docx-pdf" or "xlsx-pdf" or "pptx-pdf"
                    ? settings.MaximumXmlCharactersInPart : 0), ConfigurationContext.OfficeConversionRenderingConfiguration);
            stream.FlushFinalBlock();
        }
        return Convert.ToHexString(hash.Hash!);
    }

    private sealed class PdfConfigurationConverter : JsonConverter<PdfOptions> {
        public override PdfOptions Read(ref Utf8JsonReader reader, Type type, JsonSerializerOptions options) => throw new NotSupportedException();
        public override void Write(Utf8JsonWriter writer, PdfOptions value, JsonSerializerOptions options) {
            if (value.TextShapingProvider != null)
                throw new NotSupportedException("Runtime text shaping providers require an ordinary batch without checkpoints. Embedded font settings support checkpoints.");
            if (value.Encryption?.AesCryptographyProvider != null)
                throw new NotSupportedException("Runtime cryptography providers require an ordinary batch without checkpoints.");
            writer.WriteStartObject();
            writer.WritePropertyName("Settings");
            JsonSerializer.Serialize(writer, value, OfficeConversionPdfSettingsJsonContext.Default.PdfOptions);
            writer.WriteString("ExplicitSettings", value.CheckpointExplicitSettings);
            writer.WriteStartArray("EmbeddedStandardFonts");
            foreach (PdfEmbeddedFont font in value.CheckpointEmbeddedFonts) {
                writer.WriteStartObject();
                writer.WriteString("Slot", font.Font.ToString());
                writer.WriteString("Name", font.FontName);
                writer.WriteBoolean("SyntheticOblique", font.SyntheticOblique);
                writer.WriteString("ProgramSha256", Convert.ToHexString(SHA256.HashData(font.DataSnapshot)));
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
            writer.WriteEndObject();
        }
    }

    // Include every regular file: Markdown identifies images by bytes, not filename extensions.
    // Reject links so native resource reads cannot escape the fingerprinted tree.
    private static async Task<string> CaptureResourceIdentityAsync(string root, string input, long maximumBytes, CancellationToken token) {
        var resources = new List<(string Path, string Hash)>();
        string inputIdentity = OfficePathIdentity.GetPathIdentityKey(input);
        long bytes = 0;
        foreach (string path in Discover(root, true, token, rejectLinks: true)) {
            // The primary input has its own immutable/retryable hash; do not also make it configuration.
            if (OfficePathIdentity.GetPathIdentityKey(path) == inputIdentity) continue;
            if (resources.Count >= OfficeWorkflowHtmlResourceResolver.MaximumReferencedResourceCount)
                throw new InvalidDataException("The source resource directory exceeds the checkpoint resource limit.");
            EnsureNoLinks(path);
            using (var file = OfficePathIdentity.OpenRegularFileForRead(path, root, 81920)) bytes = checked(bytes + file.Length);
            if (bytes > maximumBytes) throw new InvalidDataException("The source resources exceed the checkpoint byte budget.");
            resources.Add((OfficePathIdentity.GetPathIdentityKey(path), await HashFileAsync(path, root, maximumBytes, token).ConfigureAwait(false)));
        }
        return Hash(string.Join("\n", resources.OrderBy(item => item.Path, StringComparer.Ordinal).Select(item => item.Path + "\0" + item.Hash)));
    }

    private static async Task VerifyResourcesAsync(string? root, string input, string? expected, long limit, CancellationToken token) {
        if (root != null && await CaptureResourceIdentityAsync(root, input, limit, token).ConfigureAwait(false) != expected)
            throw new InvalidDataException("Conversion resources changed before publication; no output was replaced.");
    }
}
