using System.Globalization;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using Org.BouncyCastle.Asn1;
using Org.BouncyCastle.Cms;
using Org.BouncyCastle.Crypto.Engines;
using Org.BouncyCastle.Crypto.Parameters;
using Org.BouncyCastle.Security;

namespace OfficeIMO.Tests.Pdf;

/// <summary>
/// Writes public-key encrypted PDFs (ISO 32000-2, 7.6.5) independently of OfficeIMO.Pdf: Bouncy Castle builds the
/// recipient envelopes and this class derives keys and encrypts objects directly from the specification.
/// </summary>
public static class PublicKeyEncryptedPdfFixture {
    internal const string Text = "Public key secret text";
    internal const string Title = "Public key title";
    internal const int AllPermissions = unchecked((int)0xFFFFFFFC);

    public enum Cipher { Rc4, AesV2, AesV3 }

    internal sealed class Settings {
        internal string SubFilter { get; init; } = "adbe.pkcs7.s5";
        internal int Version { get; init; } = 5;
        internal Cipher Method { get; init; } = Cipher.AesV3;
        internal int KeyLengthBits { get; init; } = 256;
        internal bool EncryptMetadata { get; init; } = true;
        internal int Permissions { get; init; } = AllPermissions;
        internal string EnvelopeAlgorithm { get; init; } = CmsEnvelopedGenerator.Aes256Cbc;
        internal bool SingleRecipientString { get; init; }
        internal int RecipientContentLength { get; init; } = 24;
        internal string? StringFilterOverride { get; init; }
        internal string? CfmOverride { get; init; }
        internal string? DocumentLengthOverride { get; init; }
        internal string? FilterLengthOverride { get; init; }
        internal string? EmbeddedFileFilter { get; init; }
    }

    internal static X509Certificate2 CreateRecipient(string name) {
        using RSA rsa = RSA.Create(2048);
        var request = new CertificateRequest("CN=" + name, rsa, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
        request.CertificateExtensions.Add(new X509KeyUsageExtension(X509KeyUsageFlags.KeyEncipherment, critical: true));
        using X509Certificate2 ephemeral = request.CreateSelfSigned(DateTimeOffset.UtcNow.AddMinutes(-5), DateTimeOffset.UtcNow.AddDays(1));
        return LoadPfx(ephemeral.Export(X509ContentType.Pfx, "fixture"), "fixture");
    }

    internal static X509Certificate2 LoadPfx(byte[] pfx, string password) {
#pragma warning disable SYSLIB0057 // The PKCS#12 constructor keeps the fixture loader identical on net8.0 and net10.0.
        return new X509Certificate2(pfx, password, X509KeyStorageFlags.Exportable);
#pragma warning restore SYSLIB0057
    }

    internal static byte[] Create(Settings settings, X509Certificate2 recipient, params X509Certificate2[] otherRecipients) {
        byte[] seed = new byte[20];
        using (RandomNumberGenerator random = RandomNumberGenerator.Create()) {
            random.GetBytes(seed);
        }

        byte[] recipientContent = new byte[settings.RecipientContentLength];
        Buffer.BlockCopy(seed, 0, recipientContent, 0, Math.Min(20, recipientContent.Length));
        if (recipientContent.Length >= 24) {
            recipientContent[20] = (byte)(settings.Permissions >> 24);
            recipientContent[21] = (byte)(settings.Permissions >> 16);
            recipientContent[22] = (byte)(settings.Permissions >> 8);
            recipientContent[23] = (byte)settings.Permissions;
        }

        // Acrobat writes one envelope per permission set. Put unrelated recipients first so the reader has to search.
        var envelopes = new List<byte[]>();
        foreach (X509Certificate2 other in otherRecipients) {
            envelopes.Add(Envelope(recipientContent, other, settings.EnvelopeAlgorithm));
        }

        envelopes.Add(Envelope(recipientContent, recipient, settings.EnvelopeAlgorithm));
        byte[] fileKey = DeriveFileKey(seed, envelopes, settings);

        string recipients = settings.SingleRecipientString && envelopes.Count == 1
            ? "<" + Hex(envelopes[0]) + ">"
            : "[" + string.Join(" ", envelopes.Select(envelope => "<" + Hex(envelope) + ">")) + "]";
        string encrypt = BuildEncryptionDictionary(settings, recipients);
        string metadataXml = "<?xpacket begin=\"\"?><x:xmpmeta xmlns:x=\"adobe:ns:meta/\"></x:xmpmeta><?xpacket end=\"r\"?>";
        byte[] content = Encoding.ASCII.GetBytes("BT /F1 18 Tf 72 720 Td (" + Text + ") Tj ET");
        byte[] metadata = Encoding.ASCII.GetBytes(metadataXml);

        var output = new MemoryStream();
        var offsets = new List<long>();
        Write(output, "%PDF-1.7\n");
        void Object(int number, Action body) {
            offsets.Add(output.Position);
            Write(output, number.ToString(CultureInfo.InvariantCulture) + " 0 obj\n");
            body();
            Write(output, "\nendobj\n");
        }

        void Stream(int number, string dictionary, byte[] data) {
            Object(number, () => {
                Write(output, "<< " + dictionary + " /Length " + data.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n");
                output.Write(data, 0, data.Length);
                Write(output, "\nendstream");
            });
        }

        Object(1, () => Write(output, "<< /Type /Catalog /Pages 2 0 R /Metadata 7 0 R >>"));
        Object(2, () => Write(output, "<< /Type /Pages /Kids [3 0 R] /Count 1 >>"));
        Object(3, () => Write(output, "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>"));
        Object(4, () => Write(output, "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"));
        Stream(5, string.Empty, EncryptObject(5, content, settings.Method, fileKey, settings.KeyLengthBits));
        bool stringsEncrypted = settings.StringFilterOverride != "Identity";
        byte[] title = Encoding.ASCII.GetBytes(Title);
        Object(6, () => Write(output, "<< /Title <" + Hex(stringsEncrypted ? EncryptObject(6, title, settings.Method, fileKey, settings.KeyLengthBits) : title) + "> >>"));
        Stream(7, "/Type /Metadata /Subtype /XML", settings.EncryptMetadata ? EncryptObject(7, metadata, settings.Method, fileKey, settings.KeyLengthBits) : metadata);
        Object(8, () => Write(output, encrypt));

        long xref = output.Position;
        Write(output, "xref\n0 9\n0000000000 65535 f \n");
        foreach (long offset in offsets) {
            Write(output, offset.ToString("D10", CultureInfo.InvariantCulture) + " 00000 n \n");
        }

        Write(output, "trailer\n<< /Size 9 /Root 1 0 R /Info 6 0 R /Encrypt 8 0 R /ID [<00112233445566778899aabbccddeeff> <00112233445566778899aabbccddeeff>] >>\nstartxref\n" +
            xref.ToString(CultureInfo.InvariantCulture) + "\n%%EOF\n");
        return output.ToArray();
    }

    private static string BuildEncryptionDictionary(Settings settings, string recipients) {
        var builder = new StringBuilder();
        builder.Append("<< /Filter /Adobe.PubSec /SubFilter /").Append(settings.SubFilter)
            .Append(" /V ").Append(settings.Version.ToString(CultureInfo.InvariantCulture))
            .Append(" /Length ").Append(settings.DocumentLengthOverride ?? settings.KeyLengthBits.ToString(CultureInfo.InvariantCulture));
        if (settings.SubFilter != "adbe.pkcs7.s5") {
            builder.Append(" /Recipients ").Append(recipients).Append(" >>");
            return builder.ToString();
        }

        string cfm = settings.CfmOverride ?? settings.Method switch {
            Cipher.Rc4 => "V2",
            Cipher.AesV2 => "AESV2",
            _ => "AESV3"
        };
        if (settings.EmbeddedFileFilter is not null) {
            builder.Append(" /EFF /").Append(settings.EmbeddedFileFilter);
        }

        builder.Append(" /StmF /DefaultCryptFilter /StrF /").Append(settings.StringFilterOverride ?? "DefaultCryptFilter")
            .Append(" /CF << /DefaultCryptFilter << /Type /CryptFilter /AuthEvent /DocOpen /CFM /").Append(cfm)
            .Append(" /Length ").Append(settings.FilterLengthOverride ?? settings.KeyLengthBits.ToString(CultureInfo.InvariantCulture))
            .Append(" /EncryptMetadata ").Append(settings.EncryptMetadata ? "true" : "false")
            .Append(" /Recipients ").Append(recipients)
            .Append(" >> /OtherCryptFilter << /Type /CryptFilter /CFM /AESV3 /Recipients <3000> >> >> >>");
        return builder.ToString();
    }

    private static byte[] Envelope(byte[] content, X509Certificate2 recipient, string algorithm) {
        var generator = new CmsEnvelopedDataGenerator();
        generator.AddKeyTransRecipient(DotNetUtilities.FromX509Certificate(recipient));
        return generator.Generate(new CmsProcessableByteArray(content), new DerObjectIdentifier(algorithm)).GetEncoded();
    }

    private static byte[] DeriveFileKey(byte[] seed, IReadOnlyList<byte[]> envelopes, Settings settings) {
        using HashAlgorithm hash = settings.Method == Cipher.AesV3 ? SHA256.Create() : SHA1.Create();
        var input = new MemoryStream();
        input.Write(seed, 0, seed.Length);
        foreach (byte[] envelope in envelopes) {
            input.Write(envelope, 0, envelope.Length);
        }

        if (!settings.EncryptMetadata && settings.Version >= 4) {
            input.Write(new byte[] { 0xFF, 0xFF, 0xFF, 0xFF }, 0, 4);
        }

        byte[] digest = hash.ComputeHash(input.ToArray());
        return digest.Take(settings.KeyLengthBits / 8).ToArray();
    }

    private static byte[] EncryptObject(int objectNumber, byte[] data, Cipher method, byte[] fileKey, int keyLengthBits) {
        if (method == Cipher.AesV3) {
            return AesCbc(fileKey, data);
        }

        var keyInput = new List<byte>(fileKey) {
            (byte)objectNumber, (byte)(objectNumber >> 8), (byte)(objectNumber >> 16), 0, 0
        };
        if (method == Cipher.AesV2) {
            keyInput.AddRange(new byte[] { 0x73, 0x41, 0x6C, 0x54 });
        }

        using MD5 md5 = MD5.Create();
        byte[] objectKey = md5.ComputeHash(keyInput.ToArray()).Take(Math.Min(keyLengthBits / 8 + 5, 16)).ToArray();
        if (method == Cipher.AesV2) {
            return AesCbc(objectKey, data);
        }

        var rc4 = new RC4Engine();
        rc4.Init(true, new KeyParameter(objectKey));
        byte[] output = new byte[data.Length];
        rc4.ProcessBytes(data, 0, data.Length, output, 0);
        return output;
    }

    private static byte[] AesCbc(byte[] key, byte[] data) {
        using Aes aes = Aes.Create();
        aes.Key = key;
        aes.GenerateIV();
        aes.Mode = CipherMode.CBC;
        aes.Padding = PaddingMode.PKCS7;
        using ICryptoTransform encryptor = aes.CreateEncryptor();
        byte[] ciphertext = encryptor.TransformFinalBlock(data, 0, data.Length);
        return aes.IV.Concat(ciphertext).ToArray();
    }

    private static string Hex(byte[] value) {
        var builder = new StringBuilder(value.Length * 2);
        foreach (byte item in value) {
            builder.Append(item.ToString("x2", CultureInfo.InvariantCulture));
        }

        return builder.ToString();
    }

    private static void Write(Stream stream, string text) {
        byte[] bytes = Encoding.ASCII.GetBytes(text);
        stream.Write(bytes, 0, bytes.Length);
    }
}
