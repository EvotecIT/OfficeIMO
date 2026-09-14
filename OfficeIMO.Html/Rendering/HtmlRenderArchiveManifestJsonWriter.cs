using System.Globalization;
using System.Text;

namespace OfficeIMO.Html;

/// <summary>Dependency-free deterministic JSON writer for HTML render archive manifests.</summary>
public static class HtmlRenderArchiveManifestJsonWriter {
    /// <summary>Serializes a typed archive manifest with stable ordering.</summary>
    public static string ToJson(HtmlRenderArchiveManifest manifest) {
        if (manifest == null) throw new ArgumentNullException(nameof(manifest));
        var builder = new StringBuilder();
        builder.AppendLine("{");
        String(builder, 1, "schemaId", HtmlRenderArchiveManifest.SchemaId, true);
        Number(builder, 1, "schemaVersion", HtmlRenderArchiveManifest.SchemaVersion, true);
        builder.AppendLine("  \"request\": {");
        String(builder, 2, "profileId", manifest.ProfileId, true);
        String(builder, 2, "documentState", manifest.DocumentState.ToString(), true);
        String(builder, 2, "cssMedia", manifest.CssMedia.ToString(), true);
        String(builder, 2, "surface", manifest.Surface.ToString(), true);
        String(builder, 2, "pagination", manifest.Pagination.ToString(), true);
        String(builder, 2, "pageSet", manifest.PageSet.ToString(), true);
        Number(builder, 2, "firstPageIndex", manifest.FirstPageIndex, true);
        NullableNumber(builder, 2, "pageCount", manifest.PageCount, true);
        String(builder, 2, "encoder", manifest.Encoder.ToString(), true);
        String(builder, 2, "imageFormat", manifest.ImageFormat.ToString(), true);
        String(builder, 2, "coverage", manifest.Coverage.ToString(), true);
        Number(builder, 2, "requestedScale", manifest.RequestedScale, true);
        String(builder, 2, "backgroundColor", "#" + manifest.BackgroundColor.ToHex().ToLowerInvariant(), false);
        builder.AppendLine("  },");
        Boolean(builder, 1, "hasLoss", manifest.HasLoss, true);
        StringArray(builder, 1, "providerIds", manifest.ProviderIds, true);
        builder.AppendLine("  \"pages\": [");
        for (int pageIndex = 0; pageIndex < manifest.Pages.Count; pageIndex++) {
            HtmlRenderArchivePage page = manifest.Pages[pageIndex];
            builder.AppendLine("    {");
            Number(builder, 3, "outputIndex", page.OutputIndex, true);
            Number(builder, 3, "sourcePageNumber", page.SourcePageNumber, true);
            Number(builder, 3, "width", page.Width, true);
            Number(builder, 3, "height", page.Height, true);
            Boolean(builder, 3, "isClipped", page.IsClipped, true);
            Number(builder, 3, "encodedWidth", page.EncodedWidth, true);
            Number(builder, 3, "encodedHeight", page.EncodedHeight, true);
            String(builder, 3, "mimeType", page.MimeType, true);
            Number(builder, 3, "encodedLength", page.EncodedLength, true);
            String(builder, 3, "entryName", page.EntryName, true);
            String(builder, 3, "sha256", page.Sha256, true);
            Boolean(builder, 3, "hasLoss", page.HasLoss, true);
            builder.AppendLine("      \"encodingDiagnostics\": [");
            for (int diagnosticIndex = 0; diagnosticIndex < page.EncodingDiagnostics.Count; diagnosticIndex++) {
                OfficeIMO.Drawing.OfficeImageExportDiagnostic diagnostic = page.EncodingDiagnostics[diagnosticIndex];
                builder.AppendLine("        {");
                String(builder, 5, "code", diagnostic.Code, true);
                String(builder, 5, "severity", diagnostic.Severity.ToString(), true);
                String(builder, 5, "lossKind", diagnostic.LossKind.ToString(), true);
                String(builder, 5, "message", diagnostic.Message, true);
                NullableString(builder, 5, "source", diagnostic.Source, false);
                builder.Append("        }");
                CommaLine(builder, diagnosticIndex < page.EncodingDiagnostics.Count - 1);
            }
            builder.AppendLine("      ],");
            builder.AppendLine("      \"sourcePlacements\": [");
            for (int placementIndex = 0; placementIndex < page.SourcePlacements.Count; placementIndex++) {
                HtmlRenderSourcePlacement placement = page.SourcePlacements[placementIndex];
                builder.AppendLine("        {");
                Number(builder, 5, "sourcePageNumber", placement.SourcePageNumber, true);
                Number(builder, 5, "sourceOffsetX", placement.SourceOffsetX, true);
                Number(builder, 5, "sourceOffsetY", placement.SourceOffsetY, true);
                Number(builder, 5, "width", placement.Width, true);
                Number(builder, 5, "height", placement.Height, true);
                Number(builder, 5, "outputOffsetX", placement.OutputOffsetX, true);
                Number(builder, 5, "outputOffsetY", placement.OutputOffsetY, true);
                Boolean(builder, 5, "isClipped", placement.IsClipped, false);
                builder.Append("        }");
                CommaLine(builder, placementIndex < page.SourcePlacements.Count - 1);
            }
            builder.AppendLine("      ]");
            builder.Append("    }");
            CommaLine(builder, pageIndex < manifest.Pages.Count - 1);
        }
        builder.AppendLine("  ],");
        builder.AppendLine("  \"diagnostics\": [");
        for (int index = 0; index < manifest.Diagnostics.Count; index++) {
            HtmlDiagnostic diagnostic = manifest.Diagnostics[index];
            builder.AppendLine("    {");
            String(builder, 3, "component", diagnostic.Component, true);
            String(builder, 3, "code", diagnostic.Code, true);
            String(builder, 3, "severity", diagnostic.Severity.ToString(), true);
            String(builder, 3, "lossKind", diagnostic.LossKind.ToString(), true);
            String(builder, 3, "message", diagnostic.Message, true);
            NullableString(builder, 3, "source", diagnostic.Source, true);
            NullableString(builder, 3, "detail", diagnostic.Detail, true);
            builder.AppendLine("      \"provenance\": {");
            String(builder, 4, "sourceAddress", diagnostic.Provenance.SourceAddress, true);
            Number(builder, 4, "sourceLine", diagnostic.Provenance.SourceLine, true);
            Number(builder, 4, "sourceColumn", diagnostic.Provenance.SourceColumn, true);
            String(builder, 4, "targetAddress", diagnostic.Provenance.TargetAddress, false);
            builder.AppendLine("      }");
            builder.Append("    }");
            CommaLine(builder, index < manifest.Diagnostics.Count - 1);
        }
        builder.AppendLine("  ]");
        builder.Append('}');
        return builder.ToString().Replace("\r\n", "\n").Replace('\r', '\n');
    }

    private static void StringArray(StringBuilder builder, int indent, string name,
        IReadOnlyList<string> values, bool comma) {
        Indent(builder, indent).Append('"').Append(Escape(name)).Append("\": [");
        for (int index = 0; index < values.Count; index++) {
            if (index > 0) builder.Append(", ");
            builder.Append('"').Append(Escape(values[index])).Append('"');
        }
        builder.Append(']');
        CommaLine(builder, comma);
    }

    private static void String(StringBuilder builder, int indent, string name, string value, bool comma) {
        Indent(builder, indent).Append('"').Append(Escape(name)).Append("\": \"")
            .Append(Escape(value)).Append('"');
        CommaLine(builder, comma);
    }

    private static void NullableString(StringBuilder builder, int indent, string name, string? value, bool comma) {
        if (value == null) {
            Indent(builder, indent).Append('"').Append(Escape(name)).Append("\": null");
            CommaLine(builder, comma);
        } else {
            String(builder, indent, name, value, comma);
        }
    }

    private static void NullableNumber(StringBuilder builder, int indent, string name, int? value, bool comma) {
        if (value.HasValue) {
            Number(builder, indent, name, value.Value, comma);
        } else {
            Indent(builder, indent).Append('"').Append(Escape(name)).Append("\": null");
            CommaLine(builder, comma);
        }
    }

    private static void Number(StringBuilder builder, int indent, string name, double value, bool comma) {
        Indent(builder, indent).Append('"').Append(Escape(name)).Append("\": ")
            .Append(value.ToString("R", CultureInfo.InvariantCulture));
        CommaLine(builder, comma);
    }

    private static void Number(StringBuilder builder, int indent, string name, long value, bool comma) {
        Indent(builder, indent).Append('"').Append(Escape(name)).Append("\": ")
            .Append(value.ToString(CultureInfo.InvariantCulture));
        CommaLine(builder, comma);
    }

    private static void Boolean(StringBuilder builder, int indent, string name, bool value, bool comma) {
        Indent(builder, indent).Append('"').Append(Escape(name)).Append("\": ")
            .Append(value ? "true" : "false");
        CommaLine(builder, comma);
    }

    private static StringBuilder Indent(StringBuilder builder, int count) => builder.Append(' ', count * 2);

    private static void CommaLine(StringBuilder builder, bool comma) {
        if (comma) builder.Append(',');
        builder.AppendLine();
    }

    private static string Escape(string value) {
        var builder = new StringBuilder(value.Length + 8);
        foreach (char character in value) {
            switch (character) {
                case '"': builder.Append("\\\""); break;
                case '\\': builder.Append("\\\\"); break;
                case '\b': builder.Append("\\b"); break;
                case '\f': builder.Append("\\f"); break;
                case '\n': builder.Append("\\n"); break;
                case '\r': builder.Append("\\r"); break;
                case '\t': builder.Append("\\t"); break;
                default:
                    if (character < 0x20 || character == '\u2028' || character == '\u2029') {
                        builder.Append("\\u").Append(((int)character).ToString("x4", CultureInfo.InvariantCulture));
                    } else {
                        builder.Append(character);
                    }
                    break;
            }
        }
        return builder.ToString();
    }
}
