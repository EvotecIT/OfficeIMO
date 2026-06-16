namespace OfficeIMO.Rtf.Writing;

internal static partial class RtfDocumentWriter {
    private static void WriteFontTable(StringBuilder builder, RtfDocument document, RtfWriteOptions options, int unicodeSkipCount) {
        IReadOnlyList<RtfFont> fonts = document.Fonts.Count == 0
            ? new[] { new RtfFont(0, options.DefaultFontName) }
            : document.Fonts;

        builder.Append(@"{\fonttbl");
        foreach (RtfFont font in fonts.OrderBy(font => font.Id)) {
            builder.Append(@"{\f");
            builder.Append(font.Id.ToString(CultureInfo.InvariantCulture));
            WriteFontFamily(builder, font.Family);
            AppendOptionalTwips(builder, @"\fcharset", font.Charset);
            AppendOptionalTwips(builder, @"\fprq", font.Pitch);
            AppendOptionalTwips(builder, @"\cpg", font.CodePage);
            AppendOptionalTwips(builder, @"\fbias", font.Bias);
            WriteFontDestination(builder, "panose", font.Panose, unicodeSkipCount);
            WriteFontDestination(builder, "fname", font.NonTaggedName, unicodeSkipCount);
            WriteFontEmbedding(builder, font.Embedding, unicodeSkipCount);
            builder.Append(' ');
            builder.Append(EscapeText(font.Name, unicodeSkipCount));
            WriteFontDestination(builder, "falt", font.AlternateName, unicodeSkipCount);
            builder.Append(";}");
        }

        builder.Append('}');
    }

    private static void WriteFontFamily(StringBuilder builder, RtfFontFamily? family) {
        if (!family.HasValue) return;

        builder.Append(family.Value switch {
            RtfFontFamily.Roman => @"\froman",
            RtfFontFamily.Swiss => @"\fswiss",
            RtfFontFamily.Modern => @"\fmodern",
            RtfFontFamily.Script => @"\fscript",
            RtfFontFamily.Decorative => @"\fdecor",
            RtfFontFamily.Technical => @"\ftech",
            RtfFontFamily.Bidirectional => @"\fbidi",
            _ => @"\fnil"
        });
    }

    private static void WriteFontDestination(StringBuilder builder, string destination, string? value, int unicodeSkipCount) {
        if (value == null) return;

        string trimmedValue = value.Trim();
        if (trimmedValue.Length == 0) return;

        builder.Append(@"{\*\");
        builder.Append(destination);
        builder.Append(' ');
        builder.Append(EscapeText(trimmedValue, unicodeSkipCount));
        builder.Append('}');
    }

    private static void WriteFontEmbedding(StringBuilder builder, RtfFontEmbedding? embedding, int unicodeSkipCount) {
        if (embedding == null) return;

        builder.Append(@"{\*\fontemb");
        builder.Append(embedding.Type == RtfEmbeddedFontType.TrueType ? @"\fttruetype" : @"\ftnil");
        if (!string.IsNullOrWhiteSpace(embedding.FileName) || embedding.FileCodePage.HasValue) {
            builder.Append(@"{\*\fontfile");
            AppendOptionalTwips(builder, @"\cpg", embedding.FileCodePage);
            if (!string.IsNullOrWhiteSpace(embedding.FileName)) {
                builder.Append(' ');
                builder.Append(EscapeText(embedding.FileName!.Trim(), unicodeSkipCount));
            }

            builder.Append('}');
        }

        if (embedding.Data.Length > 0) {
            builder.Append(' ');
            WriteHexBytes(builder, embedding.Data);
        }

        builder.Append('}');
    }
}
