using System.Security.Cryptography;
using System.Threading;

namespace OfficeIMO.Pdf;

/// <summary>Retains authenticated Standard security while protecting a normalized mutation output.</summary>
internal sealed class PdfSourceEncryptionContext {
    private readonly PdfStandardSecurityHandler _handler;
    private readonly PdfDictionary _dictionary;
    private readonly byte[] _permanentId;
    private readonly PdfLoadOptions _readOptions;

    internal PdfLoadOptions ReadOptions => _readOptions;

    private PdfSourceEncryptionContext(PdfStandardSecurityHandler handler, PdfDictionary dictionary,
        byte[] permanentId, PdfLoadOptions readOptions) {
        _handler = handler;
        _dictionary = dictionary;
        _permanentId = permanentId;
        _readOptions = readOptions;
    }

    internal static PdfSourceEncryptionContext? Create(PdfReadDocument document, CancellationToken cancellationToken = default) =>
        Create(document.Objects, document.TrailerRaw, document.Security, document.ReadOptions, cancellationToken);

    internal static PdfSourceEncryptionContext? Create(Dictionary<int, PdfIndirectObject> objects, string trailerRaw,
        PdfDocumentSecurityInfo security, PdfLoadOptions? readOptions, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!security.HasEncryption) return null;
        PdfLoadOptions options = PdfLoadOptions.Resolve(readOptions);
        if (!PdfSyntax.TryCreateDecryptor(objects, trailerRaw, options, out PdfStandardSecurityHandler? handler, cancellationToken) || handler is null) {
            throw new PdfUnsupportedEncryptionException("The authenticated source encryption context could not be retained.");
        }
        PdfReference reference = PdfSyntax.ReadTrailerReference(trailerRaw, "Encrypt", options.Limits, cancellationToken)
            ?? throw new PdfUnsupportedEncryptionException("The source encryption dictionary reference is missing.");
        PdfDictionary dictionary = PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect)
            ? indirect.Value as PdfDictionary ?? throw new PdfUnsupportedEncryptionException("The source encryption dictionary is unreadable.")
            : throw new PdfUnsupportedEncryptionException("The source encryption dictionary is missing.");
        EnsureDirectEncryptionDictionary(dictionary, cancellationToken);
        byte[] permanentId = PdfSyntax.ReadPermanentTrailerIdentifier(trailerRaw)
            ?? throw new PdfUnsupportedEncryptionException("The permanent file identifier required to retain encryption is missing.");
        return new PdfSourceEncryptionContext(handler, dictionary, permanentId, options);
    }

    /// <summary>Encrypts final object numbers without changing page/object mapping or password entries.</summary>
    internal byte[] Protect(byte[] plaintext, long? maximumOutputBytes = null, PdfLoadOptions? generatedReadOptions = null,
        CancellationToken cancellationToken = default) =>
        Protect(plaintext, out _, maximumOutputBytes, generatedReadOptions, cancellationToken);

    internal byte[] Protect(byte[] plaintext, out PdfGeneratedOutputGrowth growth, long? maximumOutputBytes = null,
        PdfLoadOptions? generatedReadOptions = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfLoadOptions options = PdfLoadOptions.WithMinimumInputBytes(generatedReadOptions ?? _readOptions, plaintext.LongLength);
        var (objects, trailerRaw) = PdfSyntax.ParseObjects(plaintext, options, out _, out _, cancellationToken);
        int root = PdfSyntax.ReadTrailerReference(trailerRaw, "Root", options.Limits, cancellationToken)?.ObjectNumber
            ?? throw new InvalidDataException("The mutation output does not contain a root reference.");
        int info = PdfSyntax.ReadTrailerReference(trailerRaw, "Info", options.Limits, cancellationToken)?.ObjectNumber ?? 0;
        int maximumObjectNumber = objects.Keys.Max();
        // Mutation assemblers produce classic, contiguous, generation-zero objects. Preserve their mapping.
        if (objects.Count != maximumObjectNumber || objects.Values.Any(item => item.Generation != 0 ||
            item.Value is PdfStream stream && stream.Dictionary.Get<PdfName>("Type")?.Name is "XRef" or "ObjStm")) {
            throw new NotSupportedException("Retaining encryption requires a normalized classic mutation output.");
        }
        var identityMap = objects.Keys.ToDictionary(static number => number, static number => number);
        var context = new PdfPageExtractor.SerializationContext(identityMap, 0,
            new Dictionary<int, Dictionary<string, PdfObject>>(), objects, preserveRawStringBytes: true,
            cancellationToken: cancellationToken);
        var encrypted = new List<PdfSerializedObject>(objects.Count + 1);
        int maximumCiphertextStreamBytes = 0;
        for (int number = 1; number <= maximumObjectNumber; number++) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfObject value = _handler.EncryptObject(number, 0, objects[number].Value);
            if (value is PdfStream encryptedStream) {
                maximumCiphertextStreamBytes = Math.Max(maximumCiphertextStreamBytes, encryptedStream.Data.Length);
            }
            encrypted.Add(PdfPageExtractor.SerializeIndirectObjectForAssembly(number, value, context));
        }
        int encryptionNumber = maximumObjectNumber + 1;
        encrypted.Add(PdfPageExtractor.SerializeIndirectObjectForAssembly(encryptionNumber, _dictionary, context));
        byte[] digest;
#if NET8_0_OR_GREATER
        digest = SHA256.HashData(plaintext);
#else
        using (SHA256 hash = SHA256.Create()) digest = hash.ComputeHash(plaintext);
#endif
        string trailer = " /Encrypt " + PdfSyntaxEscaper.IndirectReference(encryptionNumber) +
            " /ID [" + PdfSyntaxEscaper.HexString(_permanentId, cancellationToken) + " " + PdfSyntaxEscaper.HexString(digest.Take(16).ToArray(), cancellationToken) + "]";
        using var output = new MemoryStream();
        using var bounded = new PdfBoundedWriteStream(output, maximumOutputBytes,
            "The encrypted mutation output exceeds the configured output limit.");
        PdfFileAssembler.Assemble(bounded, encrypted, root, info,
            PdfFileAssembler.ParseHeaderVersionOrDefault(PdfSyntax.GetHeaderVersion(plaintext)), trailer, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        growth = new PdfGeneratedOutputGrowth(minimumRawStreamBytes: maximumCiphertextStreamBytes);
        return output.ToArray();
    }

    private static void EnsureDirectEncryptionDictionary(PdfObject value, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (value is PdfReference) throw new PdfUnsupportedEncryptionException("Retaining an encryption dictionary with indirect entries is not supported.");
        if (value is PdfDictionary dictionary) {
            foreach (PdfObject child in dictionary.Items.Values) EnsureDirectEncryptionDictionary(child, cancellationToken);
        } else if (value is PdfArray array) {
            foreach (PdfObject child in array.Items) EnsureDirectEncryptionDictionary(child, cancellationToken);
        }
    }
}
