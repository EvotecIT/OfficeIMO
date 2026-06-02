using System.Globalization;

namespace OfficeIMO.Pdf;

internal static class PdfOutputIntentDictionaryBuilder {
    internal static string BuildIccProfileStreamDictionary(PdfOutputIntent outputIntent, int length) {
        Guard.NotNull(outputIntent, nameof(outputIntent));
        if (length < 0) {
            throw new ArgumentOutOfRangeException(nameof(length), "PDF ICC profile stream length cannot be negative.");
        }

        return "<< /N " +
            outputIntent.ColorComponents.ToString(CultureInfo.InvariantCulture) +
            " /Length " +
            length.ToString(CultureInfo.InvariantCulture) +
            " >>";
    }

    internal static string BuildOutputIntentObject(PdfOutputIntent outputIntent, int profileObjectId) {
        Guard.NotNull(outputIntent, nameof(outputIntent));
        var sb = new StringBuilder();
        sb.Append("<< /Type /OutputIntent /S /GTS_PDFA1 /OutputConditionIdentifier ")
            .Append(PdfSyntaxEscaper.TextString(outputIntent.OutputConditionIdentifier));

        if (outputIntent.OutputCondition != null) {
            sb.Append(" /OutputCondition ")
                .Append(PdfSyntaxEscaper.TextString(outputIntent.OutputCondition));
        }

        if (outputIntent.RegistryName != null) {
            sb.Append(" /RegistryName ")
                .Append(PdfSyntaxEscaper.TextString(outputIntent.RegistryName));
        }

        if (outputIntent.Info != null) {
            sb.Append(" /Info ")
                .Append(PdfSyntaxEscaper.TextString(outputIntent.Info));
        }

        sb.Append(" /DestOutputProfile ")
            .Append(PdfSyntaxEscaper.IndirectReference(profileObjectId))
            .Append(" >>\n");
        return sb.ToString();
    }
}
