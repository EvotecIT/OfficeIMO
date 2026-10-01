using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Serialization;
using OfficeIMO.Pdf;
using OfficeIMO.Internal;

namespace OfficeIMO.Workflows;

internal static partial class OfficeConversionBatchExecutor {
    private static readonly JsonSerializerOptions ConfigurationJson = new() { IncludeFields = true,
        Converters = { new PdfConfigurationConverter() } };

    private static string CaptureConfiguration(OfficeConversionBatchRequest settings) => Hash(string.Join("\n", Schema,
        settings.InputDirectory == null ? "selected-files" : OfficePathIdentity.GetPathIdentityKey(settings.InputDirectory),
        OfficePathIdentity.GetPathIdentityKey(settings.OutputDirectory), settings.TargetExtension));

    private static string CaptureRenderingConfiguration(OfficeConversionBatchRequest settings, string routeId) {
        OfficeWorkflowConversionOptions options = settings.ConversionOptions.ForRoute(routeId);
        if (options.Html?.ResourceResolver != null || options.Html?.TextShapingProvider != null || options.Rtf?.ImageConverter != null || options.Markdown?.RemoteImageResolver != null ||
            options.Markdown?.ResourcePolicy.AllowRemoteResourceResolution == true ||
            options.Html?.ResourcePolicy.AllowRemoteResourceResolution == true)
            throw new NotSupportedException("Runtime resource callbacks and remote resources require an ordinary batch without checkpoints.");
        using var hash = SHA256.Create();
        using (var stream = new CryptoStream(Stream.Null, hash, CryptoStreamMode.Write)) {
            JsonSerializer.Serialize(stream, new { Route = routeId, settings.OutputProfile, Options = options,
                Host = Environment.MachineName }, ConfigurationJson);
            stream.FlushFinalBlock();
        }
        return Convert.ToHexString(hash.Hash!);
    }

    private sealed class PdfConfigurationConverter : JsonConverter<PdfOptions> {
        public override PdfOptions Read(ref Utf8JsonReader reader, Type type, JsonSerializerOptions options) => throw new NotSupportedException();
        public override void Write(Utf8JsonWriter writer, PdfOptions value, JsonSerializerOptions options) {
            if (value.TextShapingProvider != null)
                throw new NotSupportedException("Runtime text shaping providers require an ordinary batch without checkpoints. Embedded font settings support checkpoints.");
            writer.WriteStartObject();
            writer.WritePropertyName("Settings");
            JsonSerializer.Serialize(writer, value, new JsonSerializerOptions { IncludeFields = true });
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

    // The scoped resolver supports files under the source directory. Include their names and
    // bytes conservatively, so changed CSS, images or fonts cannot reuse an earlier artifact.
    private static async Task<string> CaptureResourceIdentityAsync(string root, long maximumBytes, CancellationToken token) {
        var resources = new List<(string Path, string Hash)>();
        long bytes = 0;
        foreach (string path in Discover(root, true, token)) {
            if (!OfficeWorkflowHtmlResourceResolver.IsSupportedDependency(path)) continue;
            if (resources.Count >= OfficeWorkflowHtmlResourceResolver.MaximumReferencedResourceCount)
                throw new InvalidDataException("The source resource directory exceeds the checkpoint resource limit.");
            EnsureNoLinks(path);
            using (var file = OfficePathIdentity.OpenRegularFileForRead(path, root, 81920)) bytes = checked(bytes + file.Length);
            if (bytes > maximumBytes) throw new InvalidDataException("The source resources exceed the checkpoint byte budget.");
            resources.Add((OfficePathIdentity.GetPathIdentityKey(path), await HashFileAsync(path, root, maximumBytes, token).ConfigureAwait(false)));
        }
        return Hash(string.Join("\n", resources.OrderBy(item => item.Path, StringComparer.Ordinal).Select(item => item.Path + "\0" + item.Hash)));
    }

    private static async Task VerifyResourcesAsync(string? root, string? expected, long limit, CancellationToken token) {
        if (root != null && await CaptureResourceIdentityAsync(root, limit, token).ConfigureAwait(false) != expected)
            throw new InvalidDataException("Conversion resources changed before publication; no output was replaced.");
    }
}
