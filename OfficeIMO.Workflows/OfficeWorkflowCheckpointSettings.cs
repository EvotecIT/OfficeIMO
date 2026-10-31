using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Serialization;
using OfficeIMO.Pdf;
using OfficeIMO.Internal;

namespace OfficeIMO.Workflows;

internal static class OfficeWorkflowCheckpointSettings {
    private static readonly JsonSerializerOptions ConfigurationJson = new() { IncludeFields = true,
        Converters = { new PdfConfigurationConverter(), new OfficeLengthFingerprintConverter(), new OfficeColorFingerprintConverter() } };
    private static readonly OfficeConversionConfigurationJsonContext ConfigurationContext = new(ConfigurationJson);

    internal static string Capture(OfficeConversionOptions settings, string routeId) {
        OfficeWorkflowRenderingOptions options = OfficeWorkflowRenderingOptions.Capture(settings).ForRoute(routeId);
        if (options.Html?.ResourceResolver != null || options.Html?.TextShapingProvider != null || options.Rtf?.ImageConverter != null ||
            options.PublisherRead?.ImageCodec != null || options.Visio?.DrawingOptions?.TextShapingProvider != null ||
            options.Visio?.DrawingOptions?.Fonts.FontProgramProvider != null || options.Visio?.DrawingOptions?.Fonts.FontVariationResolver != null ||
            options.Html?.ResourcePolicy.AllowRemoteResourceResolution == true)
            throw new NotSupportedException("Runtime resource callbacks and remote resources require an ordinary batch without checkpoints.");
        using var hash = SHA256.Create();
        using (var stream = new CryptoStream(Stream.Null, hash, CryptoStreamMode.Write)) {
            JsonSerializer.Serialize(stream, new OfficeConversionRenderingConfiguration(routeId, settings.OutputProfile, options,
                Environment.MachineName, routeId is "odg-pdf" or "visio-pdf" ? settings.MaximumXmlCharactersInPart : 0),
                ConfigurationContext.OfficeConversionRenderingConfiguration);
            stream.FlushFinalBlock();
        }
        return OfficeRuntimeCompatibility.ToHexString(hash.Hash!);
    }

    internal sealed class PdfConfigurationConverter : JsonConverter<PdfCreateOptions> {
        public override PdfCreateOptions Read(ref Utf8JsonReader reader, Type type, JsonSerializerOptions options) => throw new NotSupportedException();
        public override void Write(Utf8JsonWriter writer, PdfCreateOptions value, JsonSerializerOptions options) {
            if (value.TextShapingProvider != null)
                throw new NotSupportedException("Runtime text shaping providers require an ordinary batch without checkpoints. Embedded font settings support checkpoints.");
            if (value.Encryption?.AesCryptographyProvider != null)
                throw new NotSupportedException("Runtime cryptography providers require an ordinary batch without checkpoints.");
            writer.WriteStartObject();
            writer.WritePropertyName("Settings");
            JsonSerializer.Serialize(writer, value, OfficeBatchPdfSettingsJsonContext.Default.PdfOptions);
            writer.WriteString("ExplicitSettings", value.CheckpointExplicitSettings);
            writer.WriteStartArray("EmbeddedStandardFonts");
            foreach (PdfEmbeddedFont font in value.CheckpointEmbeddedFonts) {
                writer.WriteStartObject();
                writer.WriteString("Slot", font.Font.ToString());
                writer.WriteString("Name", font.FontName);
                writer.WriteBoolean("SyntheticOblique", font.SyntheticOblique);
                writer.WriteString("ProgramSha256", OfficeRuntimeCompatibility.ToHexString(OfficeRuntimeCompatibility.HashSha256(font.DataSnapshot)));
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
            writer.WriteEndObject();
        }
    }

}
