using DocumentFormat.OpenXml.Packaging;
using System.Threading;
using System.Threading.Tasks;
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
        ThrowIfDisposed();
        using var package = new MemoryStream();
        using (_document!.Clone(package)) { }
        package.Position = 0;
        using var snapshot = Load(package, new PowerPointLoadOptions {
            MaxInputBytes = package.Length,
            OpenSettings = new OfficeOpenXmlLoadSettings { MaxCharactersInPart = _maximumSlideXmlCharacters },
            PersistenceMode = DocumentPersistenceMode.Explicit
        });
        snapshot.SignatureMutationPolicy = SignatureMutationPolicy;
        var effective = options?.Clone() ?? new PowerPointSaveOptions();
        if (!effective.HasPasswordOverride) effective.Password = ResolveSavePassword(options);
        return snapshot.ToBytes(effective);
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
