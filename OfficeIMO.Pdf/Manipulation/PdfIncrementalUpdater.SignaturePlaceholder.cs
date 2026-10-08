using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfIncrementalUpdater {
    /// <summary>Controlled signature serialization retaining fixed-width ByteRange and cleartext CMS storage.</summary>
    internal sealed class SignaturePlaceholder {
        private SignaturePlaceholder(int objectNumber, byte[] bytes, PdfStandardSecurityHandler? encryptionHandler) {
            ObjectNumber = objectNumber;
            Bytes = bytes;
            EncryptionHandler = encryptionHandler;
        }

        internal int ObjectNumber { get; }
        internal byte[] Bytes { get; }
        internal PdfStandardSecurityHandler? EncryptionHandler { get; }

        internal static SignaturePlaceholder Create(int objectNumber, PdfExternalSignatureOptions options,
            PdfStandardSecurityHandler? encryptionHandler) => new SignaturePlaceholder(
                objectNumber,
                PdfObjectBytes.WrapIndirectObject(objectNumber, BuildSignaturePlaceholderDictionary(options, encryptionHandler, objectNumber)),
                encryptionHandler);
    }

    private static string BuildSignaturePlaceholderDictionary(PdfExternalSignatureOptions options,
        PdfStandardSecurityHandler? encryptionHandler = null, int objectNumber = 0) {
        PdfSignatureProfile profile = ResolveSignatureProfile(options);
        PdfExternalSignatureSubFilter subFilter = ResolveSignatureSubFilter(options);
        string zeros = new string('0', options.ReservedSignatureContentsBytes * 2);
        var builder = new StringBuilder();
        builder.Append("<< /Type /");
        builder.Append(profile == PdfSignatureProfile.DocumentTimestamp ? "DocTimeStamp" : "Sig");
        builder.Append(" /Filter /").Append(PdfSyntaxEscaper.Name(options.Filter));
        builder.Append(" /SubFilter /").Append(PdfSyntaxEscaper.Name(ToSubFilterName(subFilter)));
        builder.Append(" /ByteRange [").Append(SignatureByteRangePlaceholder).Append(']');
        builder.Append(" /Contents <").Append(zeros).Append('>');
        AppendSignatureTextEntry(builder, "Name", options.Name, encryptionHandler, objectNumber);
        AppendSignatureTextEntry(builder, "Reason", options.Reason, encryptionHandler, objectNumber);
        AppendSignatureTextEntry(builder, "Location", options.Location, encryptionHandler, objectNumber);
        AppendSignatureTextEntry(builder, "ContactInfo", options.ContactInfo, encryptionHandler, objectNumber);
        AppendSignatureTextEntry(builder, "M", FormatSignatureDate(options.SigningTime ?? DateTimeOffset.UtcNow), encryptionHandler, objectNumber);
        if (profile == PdfSignatureProfile.Certification) {
            builder.Append(" /Reference [<< /Type /SigRef /TransformMethod /DocMDP /TransformParams << /Type /TransformParams /P ")
                .Append(((int)options.CertificationPermission).ToString(CultureInfo.InvariantCulture))
                .Append(" /V /1.2 >> >>]");
        }
        builder.Append(" >>\n");
        return builder.ToString();
    }

    private static void AppendSignatureTextEntry(StringBuilder builder, string key, string? value,
        PdfStandardSecurityHandler? encryptionHandler, int objectNumber) {
        if (!string.IsNullOrWhiteSpace(value)) {
            string token = encryptionHandler is null
                ? PdfSyntaxEscaper.TextString(value!)
                : PdfSyntaxEscaper.HexString(((PdfStringObj)encryptionHandler.EncryptObject(
                    objectNumber, 0, new PdfStringObj(value!, useTextStringEncoding: true))).RawBytes);
            builder.Append(" /").Append(key).Append(' ').Append(token);
        }
    }

}
