using System.Text.Json;

namespace OfficeIMO.Reader;

internal static class ReaderJsonWriter {
    public static void WriteNullableString(Utf8JsonWriter writer, string name, string? value) {
        if (value == null) {
            writer.WriteNull(name);
        } else {
            writer.WriteString(name, value);
        }
    }

    public static void WriteNullableNumber(Utf8JsonWriter writer, string name, int? value) {
        if (value.HasValue) {
            writer.WriteNumber(name, value.Value);
        } else {
            writer.WriteNull(name);
        }
    }

    public static void WriteLocation(Utf8JsonWriter writer, ReaderLocation? location) {
        writer.WritePropertyName("location");
        if (location == null) {
            writer.WriteNullValue();
            return;
        }

        writer.WriteStartObject();
        WriteNullableString(writer, "path", location.Path);
        WriteNullableNumber(writer, "blockIndex", location.BlockIndex);
        WriteNullableNumber(writer, "sourceBlockIndex", location.SourceBlockIndex);
        WriteNullableNumber(writer, "startLine", location.StartLine);
        WriteNullableNumber(writer, "endLine", location.EndLine);
        WriteNullableNumber(writer, "normalizedStartLine", location.NormalizedStartLine);
        WriteNullableNumber(writer, "normalizedEndLine", location.NormalizedEndLine);
        WriteNullableString(writer, "headingPath", location.HeadingPath);
        WriteNullableString(writer, "headingSlug", location.HeadingSlug);
        WriteNullableString(writer, "sourceBlockKind", location.SourceBlockKind);
        WriteNullableString(writer, "blockAnchor", location.BlockAnchor);
        WriteNullableString(writer, "sheet", location.Sheet);
        WriteNullableString(writer, "a1Range", location.A1Range);
        WriteNullableNumber(writer, "slide", location.Slide);
        WriteNullableNumber(writer, "page", location.Page);
        WriteNullableNumber(writer, "tableIndex", location.TableIndex);
        writer.WriteEndObject();
    }
}
