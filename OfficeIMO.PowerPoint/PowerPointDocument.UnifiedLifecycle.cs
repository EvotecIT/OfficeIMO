using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.PowerPoint;

public sealed partial class PowerPointDocument {
    private string? _loadedPassword;
    private bool _savePasswordStateAssigned;
    private int? _savedLegacyKeySize;
    private bool? _savedLegacyEncryptProperties;
    private static PowerPointDocument AssociateEncryptedDocument(PowerPointDocument value, string? path, string password, Stream? source = null) {
        value._filePath = path ?? string.Empty;
        value._loadedPassword = password;
        value._savePasswordStateAssigned = true;
        if (source != null) value._sourceStream = OfficeDocumentLifecycle.ResolveAssociatedDestination(source, DocumentAccessMode.ReadWrite);
        return value;
    }
    private string? ResolveSavePassword(PowerPointSaveOptions? options) {
        string? current = _savePasswordStateAssigned ? _loadedPassword : _legacyPptPackage?.EncryptionPassword;
        return options == null ? current : options.ResolvePassword(current);
    }

    private void RecordSavedSecurity(PowerPointSaveOptions? options, bool legacy) {
        _loadedPassword = ResolveSavePassword(options);
        _savePasswordStateAssigned = true;
        if (!legacy) return;
        _savedLegacyKeySize = options?.ResolveLegacyKeySize(_savedLegacyKeySize ?? _legacyPptPackage?.EncryptionKeySizeBits)
            ?? _savedLegacyKeySize ?? _legacyPptPackage?.EncryptionKeySizeBits ?? 128;
        _savedLegacyEncryptProperties = options?.ResolveLegacyProperties(_savedLegacyEncryptProperties ?? _legacyPptPackage?.EncryptedDocumentProperties)
            ?? _savedLegacyEncryptProperties ?? _legacyPptPackage?.EncryptedDocumentProperties ?? true;
    }

    private PowerPointSaveOptions? ApplySavedLegacyProfile(PowerPointSaveOptions? options, PowerPointFileFormat format) {
        if (!IsLegacyBinaryFormat(format) || !_savedLegacyKeySize.HasValue) return options;
        if (_savedLegacyKeySize == (_legacyPptPackage?.EncryptionKeySizeBits ?? 128)
            && _savedLegacyEncryptProperties == (_legacyPptPackage?.EncryptedDocumentProperties ?? true)) return options;
        PowerPointSaveOptions effective = options?.Clone() ?? new PowerPointSaveOptions { Format = format };
        effective.LegacyPptEncryptionKeySizeBits = options?.ResolveLegacyKeySize(_savedLegacyKeySize) ?? _savedLegacyKeySize.Value;
        effective.LegacyPptEncryptDocumentProperties = options?.ResolveLegacyProperties(_savedLegacyEncryptProperties) ?? _savedLegacyEncryptProperties ?? true;
        return effective;
    }

    private byte[] CreateUnifiedBytesForPath(string path, PowerPointSaveOptions? options) {
        EnsureWritableForSave();
        PowerPointFileFormat format = PowerPointDocumentLoadRouting.GetFormat(path);
        if (IsLegacyBinaryFormat(format)) {
            PowerPointSaveOptions native = options?.Clone() ?? new PowerPointSaveOptions();
            native.Format = format;
            return ToBytes(native);
        }
        return ResolveSavePassword(options) is string password
            ? OfficeEncryption.EncryptPackage(CreateBytesForPath(path, options), password)
            : CreateBytesForPath(path, options);
    }
    /// <summary>Encodes the native presentation using one save options bag.</summary>
    /// <remarks>A read-only Open XML presentation is captured through an independent native copy. Saving the read-only document remains forbidden.</remarks>
    public byte[] ToBytes(PowerPointSaveOptions? options = null) {
        if (AccessMode == DocumentAccessMode.ReadOnly && _legacyPptPackage == null)
            return CaptureReadOnlyPackage(options);
        PowerPointFileFormat format = options?.Format ?? PowerPointFileFormat.Pptx;
        options = ApplySavedLegacyProfile(options, format);
        if (IsLegacyBinaryFormat(format) && _legacyPptPackage?.WasEncryptedSource == true) {
            if (ResolveSavePassword(options) == null)
                return CreatePlainLegacyPptBytesForSave(options);
            if (options?.HasPasswordOverride != true && options?.HasLegacyEncryptionOverride != true
                && string.Equals(ResolveSavePassword(options), _legacyPptPackage.EncryptionPassword, StringComparison.Ordinal))
                return EncodeBytes(format, options);
        }
        return ResolveSavePassword(options) is string password ? ToEncryptedBytes(password, format, options) : EncodeBytes(format, options);
    }
    private void EnsureWritableForSave() {
        ThrowIfDisposed();
        if (AccessMode == DocumentAccessMode.ReadOnly)
            throw new InvalidOperationException("The presentation is read-only and cannot be saved.");
    }

    private byte[] CaptureReadOnlyPackage(PowerPointSaveOptions? options) {
        using var package = new MemoryStream();
        CapturePackage(package);
        package.Position = 0;
        using var snapshot = Load(package, new PowerPointLoadOptions {
            MaxInputBytes = package.Length,
            OpenSettings = new OfficeOpenXmlLoadSettings { MaxCharactersInPart = _maximumSlideXmlCharacters },
            PersistenceMode = DocumentPersistenceMode.Explicit
        });
        var effective = options?.Clone() ?? new PowerPointSaveOptions();
        if (!effective.HasPasswordOverride) effective.Password = ResolveSavePassword(options);
        return snapshot.ToBytes(effective);
    }

    /// <summary>
    /// The package as it is now, edits included, copied without saving the document, so readers can work on it through
    /// the Open XML layer whichever model holds the document.
    /// </summary>
    internal byte[] CapturePackageBytes() {
        using var package = new MemoryStream();
        CapturePackage(package, processMarkupCompatibility: true);
        return package.ToArray();
    }

    private void CapturePackage(Stream package, bool processMarkupCompatibility = false) {
        ThrowIfDisposed();
        long start = package.Position;
        if (_sdkDocument == null && _sdkDeferred && _pendingEdits && _pendingPackage != null) {
            // Edits made in the package without the SDK are part of the copy, as they would be through the SDK.
            _pendingPackage.Save(package);
        } else if (_sdkDocument == null && _sdkDeferred && _packageStream is { CanRead: true, CanSeek: true } loaded) {
            // Nothing has opened the package since it was loaded, so its bytes are the copy.
            loaded.Position = 0;
            loaded.CopyTo(package);
        } else if (SdkDocument.FileOpenAccess != FileAccess.ReadWrite) {
            // A read-only SDK package keeps edits only in its loaded parts; the parts that were edited are written into the
            // copy, as edits made without the SDK are. Untouched parts keep their stored markup.
            using (var clone = (PresentationDocument)SdkDocument.Clone(package, true)) {
                if (_sdkEditedParts != null) CopyEditedParts(SdkDocument, clone, _sdkEditedParts, new HashSet<OpenXmlPart>());
                CopyPackageReferenceRelationships(SdkDocument, clone);
            }
        } else {
            // Cloning saves the SDK package first, and that save rewrites the core properties part and the core
            // properties the document holds (dates shift by the local time zone offset, and values set since the load
            // can be lost). The copy must leave the document as it was, its package fingerprints included, so the part
            // and the properties are put back. The copy itself takes the values the document held before the clone:
            // on .NET Framework the packaging layer keeps core properties in memory, and the clone doesn't carry them.
            // The clone also leaves out the package's own external relationships, which the copy needs.
            PresentationDocument document = SdkDocument;
            CorePropertyValues properties = CorePropertyValues.Read(document);
            CoreFilePropertiesPart? corePart = document.CoreFilePropertiesPart;
            byte[]? coreBytes = corePart == null ? null : ReadPartBytes(corePart);
            using (PresentationDocument copy = (PresentationDocument)document.Clone(package)) {
                if (copy.FileOpenAccess == FileAccess.ReadWrite) {
                    CopyPackageReferenceRelationships(document, copy);
                    properties.Restore(copy);
                }
            }
            if (corePart != null && coreBytes != null && !coreBytes.SequenceEqual(ReadPartBytes(corePart))) {
                using var original = new MemoryStream(coreBytes, writable: false);
                corePart.FeedData(original);
            }
            properties.Restore(document);
        }
        if (processMarkupCompatibility && _sdkDocument != null) {
            ApplyMarkupCompatibility(package, start, _sdkDocument.MarkupCompatibilityProcessSettings, _maximumSlideXmlCharacters);
        }
    }

    /// <summary>The core package property values, to put back after an SDK operation rewrites them.</summary>
    private sealed class CorePropertyValues {
        private readonly string?[] _text;
        private readonly DateTime?[] _dates;

        private CorePropertyValues(string?[] text, DateTime?[] dates) {
            _text = text;
            _dates = dates;
        }

        internal static CorePropertyValues Read(PresentationDocument document) {
            var values = document.PackageProperties;
            return new CorePropertyValues(
                new[] {
                    values.Category, values.ContentStatus, values.ContentType, values.Creator, values.Description,
                    values.Identifier, values.Keywords, values.Language, values.LastModifiedBy, values.Revision,
                    values.Subject, values.Title, values.Version
                },
                new[] { values.Created, values.Modified, values.LastPrinted });
        }

        internal void Restore(PresentationDocument document) {
            CorePropertyValues current = Read(document);
            var values = document.PackageProperties;
            if (current._text[0] != _text[0]) values.Category = _text[0];
            if (current._text[1] != _text[1]) values.ContentStatus = _text[1];
            if (current._text[2] != _text[2]) values.ContentType = _text[2];
            if (current._text[3] != _text[3]) values.Creator = _text[3];
            if (current._text[4] != _text[4]) values.Description = _text[4];
            if (current._text[5] != _text[5]) values.Identifier = _text[5];
            if (current._text[6] != _text[6]) values.Keywords = _text[6];
            if (current._text[7] != _text[7]) values.Language = _text[7];
            if (current._text[8] != _text[8]) values.LastModifiedBy = _text[8];
            if (current._text[9] != _text[9]) values.Revision = _text[9];
            if (current._text[10] != _text[10]) values.Subject = _text[10];
            if (current._text[11] != _text[11]) values.Title = _text[11];
            if (current._text[12] != _text[12]) values.Version = _text[12];
            if (!SameDate(current._dates[0], _dates[0])) values.Created = _dates[0];
            if (!SameDate(current._dates[1], _dates[1])) values.Modified = _dates[1];
            if (!SameDate(current._dates[2], _dates[2])) values.LastPrinted = _dates[2];
        }
    }

    private static byte[] ReadPartBytes(OpenXmlPart part) {
        using Stream source = part.GetStream(FileMode.Open, FileAccess.Read);
        using var copy = new MemoryStream();
        source.CopyTo(copy);
        return copy.ToArray();
    }

    private static void CopyPackageReferenceRelationships(PresentationDocument source, PresentationDocument copy) {
        foreach (ExternalRelationship relationship in source.ExternalRelationships) {
            if (copy.ExternalRelationships.All(existing => existing.Id != relationship.Id)) {
                copy.AddExternalRelationship(relationship.RelationshipType, relationship.Uri, relationship.Id);
            }
        }
        foreach (HyperlinkRelationship relationship in source.HyperlinkRelationships) {
            if (copy.HyperlinkRelationships.All(existing => existing.Id != relationship.Id)) {
                copy.AddHyperlinkRelationship(relationship.Uri, relationship.IsExternal, relationship.Id);
            }
        }
    }

    private static bool SameDate(DateTime? left, DateTime? right) =>
        left.HasValue == right.HasValue && (!left.HasValue || left.Value.Ticks == right!.Value.Ticks && left.Value.Kind == right.Value.Kind);

    // With markup compatibility processing, the SDK rewrites a part's markup when it loads the part, so a typed read sees
    // the selected branches whether or not the part was read before. The copy gets the same: it is opened with the
    // document's processing settings and every XML part is loaded, which saves the processed markup into the copy only.
    // The document and its own saves keep the markup as stored.
    private static void ApplyMarkupCompatibility(Stream package, long start,
        DocumentFormat.OpenXml.Packaging.MarkupCompatibilityProcessSettings settings, long maximumCharactersInPart) {
        if (settings.ProcessMode == DocumentFormat.OpenXml.Packaging.MarkupCompatibilityProcessMode.NoProcess) return;
        package.Position = start;
        // The part limit applies before processing drops a branch, as it does when the document reads the part.
        var openSettings = new OpenSettings { MarkupCompatibilityProcessSettings = settings, MaxCharactersInPart = maximumCharactersInPart };
        using (PresentationDocument copy = PresentationDocument.Open(package, true, openSettings)) {
            // The save that writes the processed parts must not change the core properties part.
            CoreFilePropertiesPart? corePart = copy.CoreFilePropertiesPart;
            byte[]? coreBytes = corePart == null ? null : ReadPartBytes(corePart);
            var seen = new HashSet<OpenXmlPart>();
            var pending = new Stack<OpenXmlPartContainer>();
            // The presentation's own markup, the parts the binary writer reads; package-level parts such as ribbon
            // customizations and payloads such as custom XML stay as stored.
            if (copy.PresentationPart is { } presentationPart) {
                LoadForMarkupCompatibility(presentationPart);
                pending.Push(presentationPart);
            }
            while (pending.Count > 0) {
                foreach (IdPartPair pair in pending.Pop().Parts) {
                    OpenXmlPart part = pair.OpenXmlPart;
                    if (!seen.Add(part)) continue;
                    if (!OfficeIMO.PowerPoint.LegacyPpt.Internal.LegacyPptPackageView.IsPresentationMarkup(part.ContentType)) continue;
                    LoadForMarkupCompatibility(part);
                    pending.Push(part);
                }
            }
            if (corePart != null && coreBytes != null) {
                using var original = new MemoryStream(coreBytes, writable: false);
                corePart.FeedData(original);
            }
        }
        package.Position = start;
    }

    // A part that can't be loaded (over the part limit or malformed) keeps its stored markup in the copy; the reader then
    // reports it only if it reads that part, as reading it through the document would.
    private static void LoadForMarkupCompatibility(OpenXmlPart part) {
        try {
            _ = part.RootElement;
        } catch (Exception exception) when (exception is System.Xml.XmlException or InvalidDataException
                                                or DocumentFormat.OpenXml.Packaging.OpenXmlPackageException) {
        }
    }

    // Writes the content of each edited part into the same part of the copy.
    private static void CopyEditedParts(DocumentFormat.OpenXml.Packaging.OpenXmlPartContainer source, DocumentFormat.OpenXml.Packaging.OpenXmlPartContainer target,
        HashSet<string> edited, HashSet<DocumentFormat.OpenXml.Packaging.OpenXmlPart> seen) {
        foreach (DocumentFormat.OpenXml.Packaging.IdPartPair pair in source.Parts) {
            if (!seen.Add(pair.OpenXmlPart) || !target.TryGetPartById(pair.RelationshipId, out DocumentFormat.OpenXml.Packaging.OpenXmlPart? copy)) continue;
            if (edited.Contains(pair.OpenXmlPart.Uri.OriginalString) && pair.OpenXmlPart.IsRootElementLoaded && pair.OpenXmlPart.RootElement is { } root) {
                using var content = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(root.OuterXml), writable: false);
                copy.FeedData(content);
            }
            CopyEditedParts(pair.OpenXmlPart, copy, edited, seen);
        }
    }

    /// <summary>Saves the native presentation to a caller-owned stream.</summary>
    public void Save(Stream destination, PowerPointSaveOptions? options = null) {
        EnsureWritableForSave();
        OfficeStreamWriter.WriteAllBytes(destination, ToBytes(options));
        RecordSavedSecurity(options, IsLegacyBinaryFormat(options?.Format ?? PowerPointFileFormat.Pptx));
        _discardChangesOnDispose = false;
    }
    /// <summary>Saves the native presentation asynchronously to a caller-owned stream.</summary>
    public Task SaveAsync(Stream destination, CancellationToken cancellationToken = default) =>
        SaveAsync(destination, options: null, cancellationToken);
    /// <summary>Saves the native presentation asynchronously to a caller-owned stream.</summary>
    public async Task SaveAsync(Stream destination, PowerPointSaveOptions? options, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        EnsureWritableForSave();
        byte[] bytes = ToBytes(options);
        await OfficeStreamWriter.WriteAllBytesAsync(destination, bytes, cancellationToken).ConfigureAwait(false);
        RecordSavedSecurity(options, IsLegacyBinaryFormat(options?.Format ?? PowerPointFileFormat.Pptx));
        _discardChangesOnDispose = false;
    }
    /// <summary>Saves pending changes to the associated destination using native output options.</summary>
    public void Save(PowerPointSaveOptions options) {
        if (!string.IsNullOrEmpty(_filePath)) Save(_filePath, options);
        else if (_sourceStream != null) Save(_sourceStream, options);
        else throw new InvalidOperationException("The presentation has no associated destination.");
    }
    /// <summary>Saves pending changes asynchronously to the associated destination.</summary>
    public Task SaveAsync(PowerPointSaveOptions options, CancellationToken cancellationToken = default) {
        if (!string.IsNullOrEmpty(_filePath)) return SaveAsync(_filePath, options, cancellationToken);
        if (_sourceStream != null) return SaveAsync(_sourceStream, options, cancellationToken);
        throw new InvalidOperationException("The presentation has no associated destination.");
    }
}
