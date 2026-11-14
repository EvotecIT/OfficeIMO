namespace OfficeIMO.Pdf;

internal sealed partial class PdfStandardSecurityHandler {
    internal const string EmbeddedFileNeedsPasswordMessage = "This embedded file is encrypted with the embedded-file crypt filter; open the document with its password to read it.";

    private PdfDictionary? _cryptFilters;
    private PdfCryptMethod? _embeddedFileMethod;
    private bool _hasFileKey = true;

    /// <summary>
    /// Records the crypt filters a stream can select by name (/Crypt stream filter) and the embedded-file filter
    /// (/EFF), ISO 32000-2 7.6.6.
    /// </summary>
    private PdfStandardSecurityHandler WithCryptFilters(PdfDictionary encryptionDictionary, int version) {
        if (version >= 4) {
            _cryptFilters = encryptionDictionary.Get<PdfDictionary>("CF");
            if (encryptionDictionary.Get<PdfName>("EFF") is not null) {
                _embeddedFileMethod = ResolveCryptMethod(encryptionDictionary, "EFF", version);
            }
        }

        return this;
    }

    /// <summary>
    /// A document whose strings and streams use the Identity crypt filter is readable without authentication; only its
    /// embedded files need the password (for example /AuthEvent /EFOpen). Those stay encrypted and are marked unreadable.
    /// </summary>
    private static bool CanOpenWithoutFileKey(PdfDictionary encryptionDictionary, int version, PdfCryptMethod streamMethod, PdfCryptMethod stringMethod) =>
        version >= 4 && streamMethod == PdfCryptMethod.Identity && stringMethod == PdfCryptMethod.Identity &&
        encryptionDictionary.Get<PdfName>("EFF") is not null;

    private PdfStandardSecurityHandler WithoutFileKey() {
        _hasFileKey = false;
        return this;
    }

    /// <summary>The crypt filter a rewritten stream is encrypted with, matching how <see cref="DecryptStream"/> reads it back.</summary>
    private PdfCryptMethod ResolveOutputStreamMethod(PdfDictionary dictionary) {
        if (ShouldSkipStreamData(dictionary) && !IsEmbeddedFileWithOwnFilter(dictionary)) return PdfCryptMethod.Identity;
        if (!IsEmbeddedFileWithOwnFilter(dictionary)) return _streamMethod;
        if (!_hasFileKey && _embeddedFileMethod != PdfCryptMethod.Identity) {
            throw new PdfUnsupportedEncryptionException("Embedded files encrypted with the /EFF crypt filter can only be rewritten after opening the document with its password.");
        }

        return _embeddedFileMethod!.Value;
    }

    private bool IsEmbeddedFileWithOwnFilter(PdfDictionary dictionary) =>
        _embeddedFileMethod.HasValue && dictionary.Get<PdfName>("Type")?.Name == "EmbeddedFile";

    private PdfStream DecryptStream(int objectNumber, int generation, PdfStream stream, PdfDictionary dictionary, CancellationToken cancellationToken) {
        byte[] sourceData = stream.GetData(cancellationToken);
        string? type = dictionary.Get<PdfName>("Type")?.Name;
        PdfCryptMethod method;
        if (type == "XRef" || (!_encryptMetadata && type == "Metadata")) {
            method = PdfCryptMethod.Identity;
        } else if (TryTakeCryptFilter(dictionary, out string? filterName)) {
            if (filterName is null) {
                // The crypt filter parameters are an indirect object that is not available yet; do not guess /Identity.
                return new PdfStream(dictionary, sourceData, decodingFailed: true,
                    "The stream's /Crypt filter parameters could not be resolved, so it cannot be decrypted.") {
                    HasIncompleteSyntax = stream.HasIncompleteSyntax
                };
            }

            method = ResolveNamedCryptFilter(filterName);
        } else if (type == "EmbeddedFile" && _embeddedFileMethod.HasValue) {
            method = _embeddedFileMethod.Value;
        } else {
            method = _streamMethod;
        }

        bool decodingFailed = stream.DecodingFailed;
        string? decodingError = stream.DecodingError;
        byte[] data = sourceData;
        if (method != PdfCryptMethod.Identity) {
            if (_hasFileKey) {
                data = DecryptData(objectNumber, generation, sourceData, method, cancellationToken);
            } else if (type == "EmbeddedFile") {
                decodingFailed = true;
                decodingError = EmbeddedFileNeedsPasswordMessage;
            } else {
                // Only embedded files may stay encrypted when the document opened without its password.
                throw new PdfPasswordRequiredException("Encrypted PDF requires a password.");
            }
        }

        return new PdfStream(dictionary, data, decodingFailed, decodingError) {
            HasIncompleteSyntax = stream.HasIncompleteSyntax
        };
    }

    /// <summary>
    /// Removes a leading /Crypt stream filter and returns the crypt filter it names (ISO 32000-2 7.4.10); the stream's
    /// remaining filters then decode normally. <paramref name="filterName"/> is null when the parameters are an
    /// indirect object that cannot be resolved.
    /// </summary>
    private bool TryTakeCryptFilter(PdfDictionary dictionary, out string? filterName) {
        filterName = null;
        PdfObject? filter = Resolve(dictionary.Items.TryGetValue("Filter", out PdfObject? value) ? value : null);
        PdfObject? rawParameters = dictionary.Items.TryGetValue("DecodeParms", out PdfObject? parms) ? parms : null;
        PdfObject? parameters = Resolve(rawParameters);
        PdfDictionary? cryptParameters;
        bool isCryptFilter = filter is PdfName { Name: "Crypt" } ||
            filter is PdfArray { Items.Count: > 0 } candidates && Resolve(candidates.Items[0]) is PdfName { Name: "Crypt" };
        if (isCryptFilter && IsUnresolved(rawParameters)) return true;
        if (filter is PdfName { Name: "Crypt" }) {
            PdfObject? first = parameters as PdfDictionary ?? (parameters as PdfArray)?.Items.FirstOrDefault();
            if (IsUnresolved(first)) return true;
            cryptParameters = Resolve(first) as PdfDictionary;
            dictionary.Items.Remove("Filter");
            dictionary.Items.Remove("DecodeParms");
        } else if (filter is PdfArray filters && filters.Items.Count > 0 && Resolve(filters.Items[0]) is PdfName { Name: "Crypt" }) {
            var remainingFilters = new PdfArray();
            for (int index = 1; index < filters.Items.Count; index++) remainingFilters.Items.Add(filters.Items[index]);
            dictionary.Items["Filter"] = remainingFilters;
            cryptParameters = null;
            if (parameters is PdfArray parameterArray) {
                if (parameterArray.Items.Count > 0 && IsUnresolved(parameterArray.Items[0])) return true;
                cryptParameters = parameterArray.Items.Count > 0 ? Resolve(parameterArray.Items[0]) as PdfDictionary : null;
                var remainingParameters = new PdfArray();
                for (int index = 1; index < parameterArray.Items.Count; index++) remainingParameters.Items.Add(parameterArray.Items[index]);
                dictionary.Items["DecodeParms"] = remainingParameters;
            } else if (parameters is PdfDictionary) {
                cryptParameters = (PdfDictionary)parameters;
                dictionary.Items.Remove("DecodeParms");
            }
        } else {
            return false;
        }

        filterName = cryptParameters?.Get<PdfName>("Name")?.Name ?? "Identity";
        return true;
    }

    /// <summary>Resolves indirect /Filter and /DecodeParms values while the document's objects are decrypted.</summary>
    internal Func<PdfReference, PdfObject?>? ReferenceResolver { get; set; }

    private bool IsUnresolved(PdfObject? value) => value is PdfReference && Resolve(value) is null;

    private PdfObject? Resolve(PdfObject? value) =>
        value is PdfReference reference && ReferenceResolver is not null ? ReferenceResolver(reference) : value;

    private PdfCryptMethod ResolveNamedCryptFilter(string filterName) {
        if (string.Equals(filterName, "Identity", StringComparison.Ordinal)) return PdfCryptMethod.Identity;
        PdfDictionary? filter = _cryptFilters?.Get<PdfDictionary>(filterName);
        if (filter is null) {
            throw new PdfUnsupportedEncryptionException("A stream selects the crypt filter /" + filterName + ", which the encryption dictionary does not define.");
        }

        return ResolveCryptFilterMethod(filter);
    }

    private static PdfCryptMethod ResolveCryptFilterMethod(PdfDictionary? filter) {
        // ISO 32000 table 25: an omitted /CFM means /None.
        string cfm = filter?.Get<PdfName>("CFM")?.Name ?? "None";
        switch (cfm) {
            case "None":
            case "Identity":
                return PdfCryptMethod.Identity;
            case "V2":
                return PdfCryptMethod.Rc4;
            case "AESV2":
                return PdfCryptMethod.AesV2;
            case "AESV3":
                return PdfCryptMethod.AesV3;
            case "AESV4":
                return PdfCryptMethod.AesV4;
            default:
                throw new PdfUnsupportedEncryptionException("Unsupported PDF crypt filter method /" + cfm + ".");
        }
    }
}
