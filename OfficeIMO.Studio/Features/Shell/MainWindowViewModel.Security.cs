using System.Collections.ObjectModel;
using System.Security.Cryptography.X509Certificates;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindowViewModel {
    private static readonly PdfBatesPositionChoice[] AvailableBatesPositions = {
        new(PdfBatesPosition.BottomLeft, "Bottom left"),
        new(PdfBatesPosition.BottomCenter, "Bottom center"),
        new(PdfBatesPosition.BottomRight, "Bottom right"),
        new(PdfBatesPosition.TopLeft, "Top left"),
        new(PdfBatesPosition.TopCenter, "Top center"),
        new(PdfBatesPosition.TopRight, "Top right")
    };

    [ObservableProperty]
    private string _protectUserPassword = string.Empty;

    [ObservableProperty]
    private string _protectConfirmPassword = string.Empty;

    [ObservableProperty]
    private string _protectOwnerPassword = string.Empty;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(CanChangeProtection))]
    [NotifyPropertyChangedFor(nameof(CanRemoveProtection))]
    private string _currentOwnerPassword = string.Empty;

    [ObservableProperty]
    private bool _protectAllowPrint = true;

    [ObservableProperty]
    private bool _protectAllowHighQualityPrint = true;

    [ObservableProperty]
    private bool _protectAllowCopy = true;

    [ObservableProperty]
    private bool _protectAllowAccessibility = true;

    [ObservableProperty]
    private bool _protectAllowAnnotations = true;

    [ObservableProperty]
    private bool _protectAllowFormFill = true;

    [ObservableProperty]
    private bool _protectAllowDocumentAssembly = true;

    [ObservableProperty]
    private bool _protectAllowContentChanges = true;

    [ObservableProperty]
    private bool _protectEncryptMetadata = true;

    [ObservableProperty]
    private PdfSigningCertificateViewModel? _selectedSigningCertificate;

    [ObservableProperty]
    private string _signatureFieldName = "Signature1";

    [ObservableProperty]
    private string _signatureReason = "Approved";

    [ObservableProperty]
    private string _signatureLocation = string.Empty;

    [ObservableProperty]
    private bool _signatureIsVisible = true;

    [ObservableProperty]
    private int _signaturePageNumber = 1;

    [ObservableProperty]
    private double _signatureX = 36D;

    [ObservableProperty]
    private double _signatureY = 36D;

    [ObservableProperty]
    private double _signatureWidth = 210D;

    [ObservableProperty]
    private double _signatureHeight = 54D;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasSignatureValidationSummary))]
    private string? _signatureValidationSummary;

    [ObservableProperty]
    private long _batesStartNumber = 1L;

    [ObservableProperty]
    private int _batesMinimumDigits = 6;

    [ObservableProperty]
    private string _batesPrefix = string.Empty;

    [ObservableProperty]
    private string _batesSuffix = string.Empty;

    [ObservableProperty]
    private PdfBatesPositionChoice _selectedBatesPosition = AvailableBatesPositions[2];

    public ObservableCollection<PdfSigningCertificateViewModel> SigningCertificates { get; } = new();

    public ObservableCollection<string> SignatureValidationFindings { get; } = new();

    public ObservableCollection<PdfBatesPositionChoice> BatesPositions { get; } = new(AvailableBatesPositions);

    public bool IsDocumentEncrypted => _workspace?.IsEncrypted == true;

    public bool HasDocumentSignatures => _workspace?.DocumentInfo.Security.HasSignatures == true;

    public bool CanChangeProtection => _workspace?.CanChangeEncryption(CurrentOwnerPassword) == true;

    public bool CanRemoveProtection => IsDocumentEncrypted && CanChangeProtection;

    public bool CanApplyCertificateSignature => _workspace?.CanSign == true && SelectedSigningCertificate is not null;

    public bool HasSignatureValidationSummary => !string.IsNullOrWhiteSpace(SignatureValidationSummary);

    partial void OnSelectedSigningCertificateChanged(PdfSigningCertificateViewModel? value) =>
        OnPropertyChanged(nameof(CanApplyCertificateSignature));

    [RelayCommand]
    private async Task SaveProtectedCopyAsync(CancellationToken cancellationToken) {
        if (_workspace is null) return;
        if (string.IsNullOrWhiteSpace(ProtectUserPassword)) {
            ErrorMessage = "Enter a document-open password.";
            return;
        }
        if (!string.Equals(ProtectUserPassword, ProtectConfirmPassword, StringComparison.Ordinal)) {
            ErrorMessage = "The document-open passwords do not match.";
            return;
        }
        string? path = await _pickSavePdf(cancellationToken).ConfigureAwait(true);
        if (string.IsNullOrWhiteSpace(path)) return;
        var encryption = new PdfStandardEncryptionOptions(ProtectUserPassword) {
            OwnerPassword = string.IsNullOrWhiteSpace(ProtectOwnerPassword) ? null : ProtectOwnerPassword,
            EncryptMetadata = ProtectEncryptMetadata,
            AllowedPermissions = BuildProtectionPermissions()
        };
        bool succeeded = await RunStandaloneAsync(
            token => _workspace.SaveProtectedCopyAsync(path, encryption, CurrentOwnerPassword, token, CreateProgress()),
            cancellationToken).ConfigureAwait(true);
        if (succeeded) {
            ProtectUserPassword = string.Empty;
            ProtectConfirmPassword = string.Empty;
            ProtectOwnerPassword = string.Empty;
            CurrentOwnerPassword = string.Empty;
            OperationStatus = "Protected copy saved";
        }
    }

    [RelayCommand]
    private async Task SaveDecryptedCopyAsync(CancellationToken cancellationToken) {
        if (_workspace is null || !IsDocumentEncrypted) return;
        string? path = await _pickSavePdf(cancellationToken).ConfigureAwait(true);
        if (string.IsNullOrWhiteSpace(path)) return;
        bool succeeded = await RunStandaloneAsync(
            token => _workspace.SaveDecryptedCopyAsync(path, CurrentOwnerPassword, token, CreateProgress()),
            cancellationToken).ConfigureAwait(true);
        if (succeeded) {
            CurrentOwnerPassword = string.Empty;
            OperationStatus = "Decrypted copy saved";
        }
    }

    [RelayCommand]
    private void RefreshSigningCertificates() {
        string? selectedThumbprint = SelectedSigningCertificate?.Thumbprint;
        SigningCertificates.Clear();
        try {
            using var store = new X509Store(StoreName.My, StoreLocation.CurrentUser);
            store.Open(OpenFlags.ReadOnly | OpenFlags.OpenExistingOnly);
            foreach (X509Certificate2 certificate in store.Certificates
                         .Find(X509FindType.FindByTimeValid, DateTime.Now, validOnly: false)
                         .OfType<X509Certificate2>()
                         .Where(static certificate => certificate.HasPrivateKey)
                         .OrderBy(static certificate => certificate.NotAfter)) {
                string displayName = certificate.GetNameInfo(X509NameType.SimpleName, forIssuer: false);
                if (string.IsNullOrWhiteSpace(displayName)) displayName = certificate.Subject;
                SigningCertificates.Add(new PdfSigningCertificateViewModel(
                    certificate.Thumbprint,
                    displayName,
                    certificate.NotAfter,
                    certificate.Issuer));
            }
            SelectedSigningCertificate = SigningCertificates.FirstOrDefault(certificate =>
                string.Equals(certificate.Thumbprint, selectedThumbprint, StringComparison.OrdinalIgnoreCase))
                ?? SigningCertificates.FirstOrDefault();
            OperationStatus = SigningCertificates.Count == 0
                ? "No signing certificates with private keys were found"
                : $"Found {SigningCertificates.Count} signing certificate(s)";
        } catch (Exception ex) {
            ErrorMessage = "The certificate store could not be read: " + ex.Message;
        }
    }

    [RelayCommand]
    private async Task ApplyCertificateSignatureAsync(CancellationToken cancellationToken) {
        if (_workspace is null || SelectedSigningCertificate is null) return;
        using X509Certificate2 certificate = LoadSigningCertificate(SelectedSigningCertificate.Thumbprint);
        string signerName = certificate.GetNameInfo(X509NameType.SimpleName, forIssuer: false);
        var options = new PdfExternalSignatureOptions {
            FieldName = SignatureFieldName,
            Name = string.IsNullOrWhiteSpace(signerName) ? null : signerName,
            Reason = string.IsNullOrWhiteSpace(SignatureReason) ? null : SignatureReason.Trim(),
            Location = string.IsNullOrWhiteSpace(SignatureLocation) ? null : SignatureLocation.Trim(),
            VisibleAppearance = SignatureIsVisible ? new PdfVisibleSignatureAppearanceOptions {
                PageNumber = SignaturePageNumber,
                X = SignatureX,
                Y = SignatureY,
                Width = SignatureWidth,
                Height = SignatureHeight,
                Text = string.IsNullOrWhiteSpace(signerName) ? "Digitally signed" : "Digitally signed by " + signerName
            } : null
        };
        bool succeeded = await RunMutationAsync(
            token => _workspace.SignAsync(certificate, options, token, CreateProgress()),
            cancellationToken).ConfigureAwait(true);
        if (succeeded) await ValidateSignaturesAsync(cancellationToken).ConfigureAwait(true);
    }

    [RelayCommand]
    private async Task ValidateSignaturesAsync(CancellationToken cancellationToken) {
        if (_workspace is null) return;
        PdfSignatureValidationReport? report = null;
        bool succeeded = await RunStandaloneAsync(
            async token => report = await _workspace.ValidateSignaturesAsync(token).ConfigureAwait(true),
            cancellationToken).ConfigureAwait(true);
        if (!succeeded || report is null) return;
        SignatureValidationSummary = report.HasSignatures
            ? $"{report.SignatureCount} signature(s) · {report.ProofStatus}"
            : "No certificate signatures found";
        SignatureValidationFindings.Clear();
        foreach (PdfSignatureValidationResult signature in report.Signatures) {
            string name = signature.Signature.SignerName ?? signature.Signature.FieldName ?? "Unnamed signature";
            SignatureValidationFindings.Add(name + " · " + (signature.IsStructurallyValid ? "structure valid" : "structural issue"));
        }
        foreach (PdfSignatureValidationFinding finding in report.Findings.Where(static finding => finding.Severity != PdfDiagnosticSeverity.Info)) {
            SignatureValidationFindings.Add(finding.Severity + " · " + finding.Message);
        }
        OperationStatus = "Signature validation complete";
    }

    private void ResetDocumentSecurityState() {
        ProtectUserPassword = string.Empty;
        ProtectConfirmPassword = string.Empty;
        ProtectOwnerPassword = string.Empty;
        CurrentOwnerPassword = string.Empty;
        ClearSignatureValidation();
    }

    private void ClearSignatureValidation() {
        SignatureValidationSummary = null;
        SignatureValidationFindings.Clear();
    }

    [RelayCommand]
    private async Task ApplyBatesNumberingAsync(CancellationToken cancellationToken) {
        if (_workspace is null) return;
        var options = new PdfBatesNumberingOptions {
            StartNumber = BatesStartNumber,
            MinimumDigits = BatesMinimumDigits,
            Prefix = BatesPrefix ?? string.Empty,
            Suffix = BatesSuffix ?? string.Empty,
            Position = SelectedBatesPosition.Position
        };
        await RunMutationAsync(
            token => _workspace.ApplyBatesNumberingAsync(options, token, CreateProgress()),
            cancellationToken).ConfigureAwait(true);
    }

    private PdfStandardPermissions BuildProtectionPermissions() {
        PdfStandardPermissions permissions = PdfStandardPermissions.None;
        if (ProtectAllowPrint) permissions |= PdfStandardPermissions.Print;
        if (ProtectAllowHighQualityPrint) permissions |= PdfStandardPermissions.HighQualityPrint;
        if (ProtectAllowCopy) permissions |= PdfStandardPermissions.CopyContents;
        if (ProtectAllowAccessibility) permissions |= PdfStandardPermissions.Accessibility;
        if (ProtectAllowAnnotations) permissions |= PdfStandardPermissions.ModifyAnnotations;
        if (ProtectAllowFormFill) permissions |= PdfStandardPermissions.FillForms;
        if (ProtectAllowDocumentAssembly) permissions |= PdfStandardPermissions.AssembleDocument;
        if (ProtectAllowContentChanges) permissions |= PdfStandardPermissions.ModifyContents;
        return permissions;
    }

    private static X509Certificate2 LoadSigningCertificate(string thumbprint) {
        using var store = new X509Store(StoreName.My, StoreLocation.CurrentUser);
        store.Open(OpenFlags.ReadOnly | OpenFlags.OpenExistingOnly);
        X509Certificate2? certificate = store.Certificates
            .Find(X509FindType.FindByThumbprint, thumbprint, validOnly: false)
            .OfType<X509Certificate2>()
            .FirstOrDefault(static candidate => candidate.HasPrivateKey);
        return certificate is null
            ? throw new InvalidOperationException("The selected signing certificate is no longer available with its private key.")
            : new X509Certificate2(certificate);
    }
}
