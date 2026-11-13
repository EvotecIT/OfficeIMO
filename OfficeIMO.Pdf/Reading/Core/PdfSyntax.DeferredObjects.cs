namespace OfficeIMO.Pdf;

// Objects inside object streams are validated at load and built on first use. The scan below walks the
// same tokens as ParseTopLevelObject and ParseObject. It accepts an object only when the parser would
// return a complete dictionary or array, and then records what the load-time passes need from it.
// Anything else - incomplete syntax, duplicate keys, escaped or non-ASCII names, limit failures - is
// left to the eager parser, so its result and diagnostics stay exactly as before.
internal static partial class PdfSyntax {
    private const int MaximumScannedKeysPerDictionary = 64;

    internal static PdfObject ParseDeferredObject(byte[] data, int start, int length, PdfReadLimits limits) =>
        ParseTopLevelObject(data, start, length, limits, trackEncodedStringSourceSpans: false, CancellationToken.None)
        ?? throw new InvalidOperationException("A deferred PDF object no longer parses as it did when it was validated.");

    internal static PdfObject ParseDeferredClassicDictionary(byte[] data, int start, int length, PdfReadLimits limits) =>
        ParseDictionary(new PdfByteContentText(data), start, length, limits, cancellationToken: CancellationToken.None);

    // A top-level indirect object of the input file whose value is a dictionary that the eager load would store
    // as is: complete syntax, no stream and nothing but whitespace or comments before endobj. The summary spans
    // the text between << and >>, the same text the eager load passes to ParseDictionary.
    private static bool TryScanDeferrableClassicDictionary(
        PdfByteContentText text,
        int dictionaryStart,
        int dictionaryEnd,
        int objectEnd,
        int bodyEnd,
        PdfReadLimits limits,
        DeferredObjectScan scan,
        out PdfDeferredObjectSummary summary,
        CancellationToken cancellationToken) {
        summary = null!;
        if (IndexOfKeywordCancellable(text, "stream", dictionaryEnd, objectEnd, cancellationToken) >= 0 ||
            SkipWhitespaceAndComments(text, dictionaryEnd, bodyEnd, cancellationToken) < bodyEnd) return false;
        int contentStart = dictionaryStart + 2;
        int contentLength = dictionaryEnd - contentStart;
        scan.Reset();
        try {
            using var tokens = Tokenize(text, contentStart, contentLength, limits, true, cancellationToken);
            if (!ScanTopLevelDictionary(tokens, limits, scan, cancellationToken)) return false;
        } catch (Exception exception) when (exception is not OperationCanceledException and not OutOfMemoryException) {
            // The eager parser reproduces the same failure, or the same limit exception, for this object.
            return false;
        }

        summary = scan.CreateSummary(contentStart, contentLength);
        return true;
    }

    private static bool TryScanDeferrableObject(
        byte[] data,
        int start,
        int length,
        PdfReadLimits limits,
        DeferredObjectScan scan,
        out PdfDeferredObjectSummary summary,
        CancellationToken cancellationToken) {
        summary = null!;
        if (length > limits.MaxObjectCharacters) return false;
        var text = new PdfByteContentText(data);
        int end = start + length;
        int first = start;
        while (first < end && char.IsWhiteSpace(text[first])) first++;
        int trimmedLength = end - first;
        if (trimmedLength < 2) return false;
        scan.Reset();
        try {
            if (text[first] == '<' && text[first + 1] == '<') {
                int dictEnd = FindDictEnd(text, first, end, cancellationToken);
                if (dictEnd <= first || SkipWhitespaceAndComments(text, dictEnd, end, cancellationToken) < end) return false;
                int dictionaryCharacters = dictEnd - (first + 2);
                if (dictionaryCharacters > limits.MaxObjectCharacters) return false;
                using var tokens = Tokenize(text, first + 2, dictionaryCharacters, limits, false, cancellationToken);
                if (!ScanTopLevelDictionary(tokens, limits, scan, cancellationToken)) return false;
            } else if (text[first] == '[') {
                using var tokens = Tokenize(text, first, trimmedLength, limits, false, cancellationToken);
                if (!ScanObject(tokens, 0, limits, 0, scan, TopLevelKey.None, out int consumed, cancellationToken) ||
                    consumed + 1 != tokens.Count) return false;
            } else {
                // Scalars are cheap to build and are left to the eager parser.
                return false;
            }
        } catch (Exception exception) when (exception is not OperationCanceledException and not OutOfMemoryException) {
            // The eager parser reproduces the same failure, or the same limit exception, for this object.
            return false;
        }

        summary = scan.CreateSummary(start, length);
        return true;
    }

    // Mirrors ParseDictionary over the tokens between the outer << and its closing >>.
    private static bool ScanTopLevelDictionary(
        in PooledTokenBuffer<PdfByteContentText> tokens,
        PdfReadLimits limits,
        DeferredObjectScan scan,
        CancellationToken cancellationToken) {
        PdfByteContentText source = tokens.Source;
        int keysStart = scan.BeginDictionary();
        for (int i = 0; i < tokens.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfToken token = tokens[i];
            if (i == 0 && token.Equals(source, "<<")) continue;
            if (token.Length > 0 && token.GetCharacter(source, 0) == '/') {
                if (!scan.TryAddKey(token, source, keysStart)) return false;
                if (i + 1 >= tokens.Count || tokens[i + 1].Equals(source, ">>")) return false;
                TopLevelKey key = ClassifyTopLevelKey(token, source);
                if (key == TopLevelKey.Encrypt) scan.MarkEncryptKey();
                if (!ScanObject(tokens, i + 1, limits, 0, scan, key, out int consumed, cancellationToken)) return false;
                i += consumed + 1;
            } else if (token.Equals(source, ">>") && i == tokens.Count - 1) {
                break;
            } else {
                return false;
            }
        }
        scan.EndDictionary(keysStart);
        return true;
    }

    // Mirrors ParseObject. Returns false wherever the parser would mark the value incomplete.
    private static bool ScanObject(
        in PooledTokenBuffer<PdfByteContentText> tokens,
        int i,
        PdfReadLimits limits,
        int depth,
        DeferredObjectScan scan,
        TopLevelKey topLevelKey,
        out int consumed,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        consumed = 0;
        if (i < 0 || i >= tokens.Count) return false;
        if (depth > limits.MaxObjectNestingDepth) return false;

        PdfToken token = tokens[i];
        PdfByteContentText source = tokens.Source;
        if (token.Equals(source, "<<")) {
            int keysStart = scan.BeginDictionary();
            int j = i + 1;
            while (j < tokens.Count && !tokens[j].Equals(source, ">>")) {
                cancellationToken.ThrowIfCancellationRequested();
                PdfToken keyToken = tokens[j];
                if (keyToken.Length == 0 || keyToken.GetCharacter(source, 0) != '/' ||
                    j + 1 >= tokens.Count || tokens[j + 1].Equals(source, ">>") ||
                    !scan.TryAddKey(keyToken, source, keysStart) ||
                    !ScanObject(tokens, j + 1, limits, depth + 1, scan, TopLevelKey.None, out int used, cancellationToken)) {
                    return false;
                }
                j += used + 2;
            }
            if (j >= tokens.Count) return false;
            scan.EndDictionary(keysStart);
            consumed = j - i;
            return true;
        }
        if (token.Equals(source, "[")) {
            int j = i + 1;
            while (j < tokens.Count && !tokens[j].Equals(source, "]")) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!ScanObject(tokens, j, limits, depth + 1, scan, TopLevelKey.None, out int used, cancellationToken)) return false;
                j += used + 1;
            }
            if (j >= tokens.Count) return false;
            consumed = j - i;
            return true;
        }
        if (token.Length > 0 && token.GetCharacter(source, 0) == '/') {
            return scan.TryAddName(token, source, topLevelKey);
        }
        if (token.IsString && token.Length > 0 && (token.GetCharacter(source, 0) == '(' ||
            token.GetCharacter(source, 0) == '<' && (token.Length == 1 || token.GetCharacter(source, 1) != '<'))) {
            return token.IsTerminated;
        }
        if (token.Equals(source, "true") || token.Equals(source, "false") || token.Equals(source, "null")) return true;
        if (token.Length > 0 && IsPdfNumberStart(token.GetCharacter(source, 0))) {
            if (i + 2 < tokens.Count && tokens[i + 2].Equals(source, "R") &&
                token.TryParseInt32(source, out int objectNumber, cancellationToken) &&
                tokens[i + 1].TryParseInt32(source, out int generation, cancellationToken)) {
                scan.AddReference(objectNumber, generation, topLevelKey);
                consumed = 2;
                return true;
            }
            return token.TryParseDouble(source, out _, cancellationToken);
        }
        return false;
    }

    private static TopLevelKey ClassifyTopLevelKey(PdfToken key, PdfByteContentText source) {
        if (key.Equals(source, "/Type")) return TopLevelKey.Type;
        if (key.Equals(source, "/Subtype")) return TopLevelKey.Subtype;
        if (key.Equals(source, "/FT")) return TopLevelKey.FieldType;
        if (key.Equals(source, "/Encrypt")) return TopLevelKey.Encrypt;
        return TopLevelKey.Other;
    }

    private enum TopLevelKey {
        None,
        Other,
        Type,
        Subtype,
        FieldType,
        Encrypt
    }

    /// <summary>Reusable state for scanning the objects of one load.</summary>
    private sealed class DeferredObjectScan {
        private readonly List<long> _references = new();
        private readonly List<PdfToken> _keys = new();
        // Types and subtypes that whole-document passes look for across every object.
        private static readonly string[] KnownTypes = { "Page", "Pages", "Annot", "Outlines", "Filespec", "Catalog", "XRef", "ObjStm", "Sig" };
        private static readonly string[] KnownSubtypes = { "Widget", "Link", "FileAttachment" };
        private string? _knownType;
        private string? _knownSubtype;
        private bool _hasEncryptKey;
        private bool _mayCarrySignatureData;

        internal void Reset() {
            _references.Clear();
            _keys.Clear();
            _knownType = null;
            _knownSubtype = null;
            _hasEncryptKey = false;
            _mayCarrySignatureData = false;
        }

        internal int BeginDictionary() => _keys.Count;

        internal void MarkEncryptKey() => _hasEncryptKey = true;

        internal void EndDictionary(int keysStart) => _keys.RemoveRange(keysStart, _keys.Count - keysStart);

        // A repeated key marks the parsed dictionary incomplete, so it is not deferred.
        internal bool TryAddKey(PdfToken key, PdfByteContentText source, int keysStart) {
            if (!IsPlainName(key, source)) return false;
            if (_keys.Count - keysStart >= MaximumScannedKeysPerDictionary) return false;
            for (int index = keysStart; index < _keys.Count; index++) {
                if (SameToken(_keys[index], key, source)) return false;
            }
            _keys.Add(key);
            NoteSignatureName(key, source);
            return true;
        }

        internal bool TryAddName(PdfToken name, PdfByteContentText source, TopLevelKey topLevelKey) {
            if (!IsPlainName(name, source)) return false;
            NoteSignatureName(name, source);
            if (topLevelKey == TopLevelKey.Type) _knownType = FindKnownName(name, source, KnownTypes);
            else if (topLevelKey == TopLevelKey.Subtype) _knownSubtype = FindKnownName(name, source, KnownSubtypes);
            return true;
        }

        private static string? FindKnownName(PdfToken name, PdfByteContentText source, string[] candidates) {
            for (int index = 0; index < candidates.Length; index++) {
                string candidate = candidates[index];
                if (name.Length != candidate.Length + 1) continue;
                bool matches = true;
                for (int character = 0; character < candidate.Length && matches; character++) {
                    matches = name.GetCharacter(source, character + 1) == candidate[character];
                }
                if (matches) return candidate;
            }
            return null;
        }

        internal void AddReference(int objectNumber, int generation, TopLevelKey topLevelKey) {
            _references.Add(((long)objectNumber << 32) | (uint)generation);
            // Signature fields and values are found by resolving /FT and /Type, which may be indirect.
            if (topLevelKey is TopLevelKey.FieldType or TopLevelKey.Type) _mayCarrySignatureData = true;
        }

        internal PdfDeferredObjectSummary CreateSummary(int start, int length) => new(
            start,
            length,
            _references.Count == 0 ? Array.Empty<long>() : _references.ToArray(),
            _knownType,
            _knownSubtype,
            _hasEncryptKey,
            _mayCarrySignatureData);

        private void NoteSignatureName(PdfToken name, PdfByteContentText source) {
            if (name.Equals(source, "/ByteRange") || name.Equals(source, "/Sig") || name.Equals(source, "/SigFlags")) {
                _mayCarrySignatureData = true;
            }
        }

        // Without # escapes or bytes above 0x7F, a name decodes to its own characters, so tokens compare as names.
        private static bool IsPlainName(PdfToken name, PdfByteContentText source) {
            for (int index = 1; index < name.Length; index++) {
                char character = name.GetCharacter(source, index);
                if (character == '#' || character >= 0x80) return false;
            }
            return true;
        }

        private static bool SameToken(PdfToken left, PdfToken right, PdfByteContentText source) {
            if (left.Length != right.Length) return false;
            for (int index = 0; index < left.Length; index++) {
                if (left.GetCharacter(source, index) != right.GetCharacter(source, index)) return false;
            }
            return true;
        }
    }
}
