using System.Globalization;
using System.Security.Cryptography;

namespace OfficeIMO.Pdf;

internal static class PdfDirectStreamIdentity {
    internal static int Compute(PdfStream stream) {
        var canonical = new StringBuilder();
        AppendObject(canonical, stream.Dictionary);
        byte[] metadata = PdfEncoding.Latin1GetBytes(canonical.ToString());
        using SHA256 sha256 = SHA256.Create();
        sha256.TransformBlock(metadata, 0, metadata.Length, metadata, 0);
        sha256.TransformFinalBlock(stream.Data, 0, stream.Data.Length);
        byte[] digest = sha256.Hash!;
        int identity = digest[0] | digest[1] << 8 | digest[2] << 16 | digest[3] << 24;
        return identity == 0 ? 1 : identity;
    }

    private static void AppendObject(StringBuilder builder, PdfObject value) {
        switch (value) {
            case PdfNumber number:
                builder.Append('n').Append(number.Value.ToString("R", CultureInfo.InvariantCulture)).Append(';');
                break;
            case PdfBoolean boolean:
                builder.Append(boolean.Value ? "b1;" : "b0;");
                break;
            case PdfName name:
                builder.Append('N').Append(name.Name.Length).Append(':').Append(name.Name).Append(';');
                break;
            case PdfStringObj text:
                builder.Append('s').Append(Convert.ToBase64String(text.RawBytes)).Append(';');
                break;
            case PdfNull:
                builder.Append("null;");
                break;
            case PdfReference reference:
                builder.Append('r').Append(reference.ObjectNumber).Append(':').Append(reference.Generation).Append(';');
                break;
            case PdfArray array:
                builder.Append('[');
                for (int index = 0; index < array.Items.Count; index++) AppendObject(builder, array.Items[index]);
                builder.Append(']');
                break;
            case PdfDictionary dictionary:
                builder.Append('{');
                foreach (KeyValuePair<string, PdfObject> entry in dictionary.Items.OrderBy(static item => item.Key, StringComparer.Ordinal)) {
                    builder.Append(entry.Key.Length).Append(':').Append(entry.Key).Append('=');
                    AppendObject(builder, entry.Value);
                }
                builder.Append('}');
                break;
            case PdfStream nestedStream:
                builder.Append("stream:").Append(Compute(nestedStream)).Append(';');
                break;
            default:
                builder.Append(value.GetType().FullName).Append(';');
                break;
        }
    }
}
