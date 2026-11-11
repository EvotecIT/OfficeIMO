using System.Globalization;
using System.Runtime.CompilerServices;

namespace OfficeIMO.Pdf;

/// <summary>
/// Converts a parsed PDF into PDF/A-2 or PDF/A-3 by fixing what does not change the rendered page and refusing,
/// with every reason, when something cannot be fixed that way.
/// </summary>
internal sealed partial class PdfAConverter {
    private readonly PdfReadDocument _document;
    private readonly Dictionary<int, PdfIndirectObject> _objects;
    private readonly PdfAConversionOptions _options;
    private readonly CancellationToken _cancellationToken;
    private readonly int _part;
    private readonly bool _unicode;
    private readonly List<PdfAConversionChange> _changes = new();
    private readonly List<PdfAConversionViolation> _violations = new();
    private readonly Dictionary<PdfObject, int> _objectNumbers = new(ReferenceComparer.Instance);
    private int _nextObjectNumber;

    private PdfAConverter(PdfReadDocument document, PdfAConversionOptions options, CancellationToken cancellationToken) {
        _document = document;
        _objects = document.Objects;
        _options = options;
        _cancellationToken = cancellationToken;
        _part = options.Profile is PdfComplianceProfile.PdfA2B or PdfComplianceProfile.PdfA2U ? 2 : 3;
        _unicode = options.Profile is PdfComplianceProfile.PdfA2U or PdfComplianceProfile.PdfA3U;
        foreach (PdfIndirectObject indirect in _objects.Values) {
            if (!_objectNumbers.ContainsKey(indirect.Value)) _objectNumbers.Add(indirect.Value, indirect.ObjectNumber);
            if (indirect.Value is PdfStream stream && !_objectNumbers.ContainsKey(stream.Dictionary)) {
                _objectNumbers.Add(stream.Dictionary, indirect.ObjectNumber);
            }
            _nextObjectNumber = Math.Max(_nextObjectNumber, indirect.ObjectNumber);
        }
        _nextObjectNumber++;
    }

    internal static PdfAConversionResult Convert(byte[] source, PdfLoadOptions? readOptions, PdfAConversionOptions options, CancellationToken cancellationToken) {
        Guard.NotNull(source, nameof(source));
        Guard.NotNull(options, nameof(options));
        if (options.Profile is not (PdfComplianceProfile.PdfA2B or PdfComplianceProfile.PdfA2U or PdfComplianceProfile.PdfA3B or PdfComplianceProfile.PdfA3U)) {
            throw new ArgumentException("PDF/A conversion supports PdfA2B, PdfA2U, PdfA3B and PdfA3U.", nameof(options));
        }
        if (options.OutputIntentProfile is { } profile && !PdfAIccProfile.TryRead(profile, out _, out string? profileError)) {
            throw new ArgumentException("The output intent profile is not a usable ICC profile: " + profileError, nameof(options));
        }

        PdfLoadOptions effectiveReadOptions = options.Password is null
            ? PdfLoadOptions.Resolve(readOptions)
            : PdfLoadOptions.WithPassword(readOptions, options.Password);
        PdfReadDocument document;
        try {
            PdfDocumentSecurityInfo security = PdfSyntax.ReadDocumentSecurityInfo(source, effectiveReadOptions, cancellationToken: cancellationToken);
            if (security.HasEncryption && !security.HasOwnerAuthorization) {
                return Refuse(options.Profile, PdfAConversionRefusalReason.EncryptionNotAuthorized, "document",
                    "The PDF is encrypted. PDF/A forbids encryption, and removing it requires the owner password in PdfAConversionOptions.Password.",
                    security.EncryptObjectNumber);
            }
            document = PdfReadDocument.Open(source, effectiveReadOptions, cancellationToken);
        } catch (Exception exception) when (IsReadFailure(exception, cancellationToken)) {
            return Refuse(options.Profile, PdfAConversionRefusalReason.SourceUnreadable, "document",
                "The PDF could not be read: " + exception.Message, null);
        }

        if (document.Security.SignatureValueCount > 0) {
            return Refuse(options.Profile, PdfAConversionRefusalReason.SignedDocument, "document",
                "The PDF has " + document.Security.SignatureValueCount.ToString(CultureInfo.InvariantCulture) +
                " signature value(s). Conversion rewrites the file, which would invalidate them. Convert the unsigned original and sign the PDF/A file.",
                document.Security.SignatureFieldObjectNumbers.Count > 0 ? document.Security.SignatureFieldObjectNumbers[0] : null);
        }

        var converter = new PdfAConverter(document, options, cancellationToken);
        int? infoObjectNumber;
        try {
            infoObjectNumber = converter.Run();
        } catch (Exception exception) when (IsReadFailure(exception, cancellationToken)) {
            converter.AddViolation(PdfAConversionRefusalReason.SourceUnreadable, "document", "The PDF could not be read completely: " + exception.Message, (int?)null);
            infoObjectNumber = null;
        }
        if (converter._violations.Count > 0) {
            return new PdfAConversionResult(null, new PdfAConversionReport(options.Profile,
                Array.Empty<PdfAConversionChange>(), converter._violations.AsReadOnly()));
        }

        if (document.Security.HasEncryption) {
            converter._changes.Insert(0, new PdfAConversionChange(PdfAConversionChangeKind.EncryptionRemoved, "document",
                "Removed password encryption, which PDF/A forbids.", document.Security.EncryptObjectNumber, OfficeConversionLossKind.None));
        }

        byte[] converted = PdfDocumentObjectGraphRewriter.Rewrite(
            source,
            document,
            effectiveReadOptions,
            outputEncryption: null,
            (_, _) => infoObjectNumber,
            preserveSourceEncryption: false,
            cancellationToken: cancellationToken,
            fileVersionOverride: PdfFileVersion.Pdf17);
        cancellationToken.ThrowIfCancellationRequested();

        string? verificationError = Verify(converted, options, cancellationToken);
        if (verificationError is not null) {
            return Refuse(options.Profile, PdfAConversionRefusalReason.VerificationFailed, "document", verificationError, null);
        }

        return new PdfAConversionResult(converted, new PdfAConversionReport(options.Profile,
            converter._changes.AsReadOnly(), Array.Empty<PdfAConversionViolation>()));
    }

    private static PdfAConversionResult Refuse(PdfComplianceProfile profile, PdfAConversionRefusalReason reason, string location, string message, int? objectNumber) =>
        new(null, new PdfAConversionReport(profile, Array.Empty<PdfAConversionChange>(),
            new[] { new PdfAConversionViolation(reason, location, message, objectNumber) }));

    private static bool IsReadFailure(Exception exception, CancellationToken cancellationToken) =>
        !(exception is OperationCanceledException && cancellationToken.IsCancellationRequested) &&
        exception is PdfParseException or InvalidDataException or PdfReadLimitException or NotSupportedException or
            FormatException or IndexOutOfRangeException or ArgumentException or InvalidOperationException or OverflowException or
            System.Security.Cryptography.CryptographicException;

    /// <summary>Reopens the written file and rechecks the conversion rules on it, so a serializer defect cannot pass silently.</summary>
    private static string? Verify(byte[] converted, PdfAConversionOptions options, CancellationToken cancellationToken) {
        PdfReadDocument output;
        try {
            output = PdfReadDocument.Open(converted, PdfLoadOptions.WithMinimumInputBytes(null, converted.LongLength), cancellationToken);
        } catch (Exception exception) when (IsReadFailure(exception, cancellationToken)) {
            return "The converted file could not be reopened: " + exception.Message;
        }
        if (output.Security.HasEncryption) return "The converted file is still encrypted.";
        var check = new PdfAConverter(output, options, cancellationToken) { _verifyOnly = true };
        check.Run();
        if (check._violations.Count == 0 && check._changes.Count == 0) return null;
        string first = check._violations.Count > 0 ? check._violations[0].ToString() : check._changes[0].ToString();
        return "The converted file still needs " + (check._violations.Count + check._changes.Count).ToString(CultureInfo.InvariantCulture) +
            " change(s); first: " + first;
    }

    private bool _verifyOnly;
    private PdfDictionary? _acroForm;
    private bool _needAppearances;

    /// <summary>Runs every fix and check; returns the Info object number to write, if any.</summary>
    private int? Run() {
        PdfDictionary catalog = _document.CatalogDictionary ?? throw new InvalidDataException("The PDF has no readable catalog.");
        RepairPageTree(catalog);
        ProcessCatalog(catalog);
        ProcessPages();
        ScanObjects(catalog);
        ProcessEmbeddedFiles(catalog);
        AnalyzeContent();
        ProcessFonts();
        ProcessColor(catalog);
        return ProcessMetadata(catalog);
    }

    internal PdfObject? Resolve(PdfObject? value) => PdfObjectLookup.ResolveChainCancellable(_objects, value, _cancellationToken);

    internal PdfDictionary? ResolveDictionary(PdfObject? value) => Resolve(value) switch {
        PdfDictionary dictionary => dictionary,
        PdfStream stream => stream.Dictionary,
        _ => null
    };

    private int? ObjectNumberOf(PdfObject? value) {
        if (value is null) return null;
        if (value is PdfReference reference) return reference.ObjectNumber;
        return _objectNumbers.TryGetValue(value, out int number) ? number : null;
    }

    private PdfReference AddObject(PdfObject value) {
        int number = _nextObjectNumber++;
        _objects[number] = new PdfIndirectObject(number, 0, value);
        _objectNumbers[value] = number;
        return new PdfReference(number, 0);
    }

    /// <summary>Records a fix. In verification mode a needed fix means the written file is not yet compliant.</summary>
    private bool Change(PdfAConversionChangeKind kind, string location, string message, PdfObject? owner,
        OfficeConversionLossKind lossKind = OfficeConversionLossKind.None) {
        _changes.Add(new PdfAConversionChange(kind, location, message, ObjectNumberOf(owner), lossKind));
        return !_verifyOnly;
    }

    private void AddViolation(PdfAConversionRefusalReason reason, string location, string message, PdfObject? owner) =>
        AddViolation(reason, location, message, ObjectNumberOf(owner));

    private void AddViolation(PdfAConversionRefusalReason reason, string location, string message, int? objectNumber) {
        foreach (PdfAConversionViolation existing in _violations) {
            if (existing.Reason == reason && existing.ObjectNumber == objectNumber && objectNumber.HasValue &&
                string.Equals(existing.Message, message, StringComparison.Ordinal)) return;
        }
        _violations.Add(new PdfAConversionViolation(reason, location, message, objectNumber));
    }

    private static string? NameOf(PdfObject? value) => (value as PdfName)?.Name;

    private string? GetName(PdfDictionary dictionary, string key) =>
        dictionary.Items.TryGetValue(key, out PdfObject? value) ? NameOf(Resolve(value)) : null;

    private static string DescribeObject(int? objectNumber) =>
        objectNumber.HasValue ? " (object " + objectNumber.Value.ToString(CultureInfo.InvariantCulture) + ")" : string.Empty;

    private sealed class ReferenceComparer : IEqualityComparer<PdfObject> {
        internal static readonly ReferenceComparer Instance = new();
        public bool Equals(PdfObject? x, PdfObject? y) => ReferenceEquals(x, y);
        public int GetHashCode(PdfObject obj) => RuntimeHelpers.GetHashCode(obj);
    }
}
