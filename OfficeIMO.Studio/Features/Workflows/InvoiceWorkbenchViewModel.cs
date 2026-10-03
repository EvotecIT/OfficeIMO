using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Internal;
using OfficeIMO.Invoicing;
using OfficeIMO.Invoicing.Pdf;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Localization;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Workflows;

/// <summary>Thin invoice workbench over shared execution, storage publication and recovery.</summary>
public sealed partial class InvoiceWorkbenchViewModel : ObservableObject, IDisposable {
    private readonly Func<CancellationToken, Task<string?>> _pickInvoice, _pickFolder;
    private readonly IOfficeInvoiceWorkflowRunner _runner;
    private readonly IStudioLocalizer _localizer;
    private readonly StudioStorageAccess? _storage;
    private readonly IOfficeWorkflowPublicationGuard? _guard;
    private readonly StudioJobHistory? _jobs;
    private readonly OfficeWorkflowOutputRecoveryStore? _recovery;
    private readonly Func<string, Task<bool>> _confirmProviderWrite;
    private readonly Func<string, CancellationToken, Task>? _openOutput;
    private CancellationTokenSource? _cancellation;
    private bool _disposed;

    /// <summary>Creates a workbench with host file pickers and the shared invoice execution owner.</summary>
    public InvoiceWorkbenchViewModel(Func<CancellationToken, Task<string?>> pickInvoice,
        Func<CancellationToken, Task<string?>> pickFolder, IOfficeInvoiceWorkflowRunner? runner = null)
        : this(pickInvoice, pickFolder, runner, null) { }

    internal InvoiceWorkbenchViewModel(Func<CancellationToken, Task<string?>> pickInvoice,
        Func<CancellationToken, Task<string?>> pickFolder, IOfficeInvoiceWorkflowRunner? runner,
        IStudioLocalizer? localizer, StudioStorageAccess? storage = null, IOfficeWorkflowPublicationGuard? guard = null,
        StudioJobHistory? jobs = null, OfficeWorkflowOutputRecoveryStore? recovery = null,
        Func<string, Task<bool>>? confirmProviderWrite = null, Func<string, CancellationToken, Task>? openOutput = null) {
        _pickInvoice = pickInvoice; _pickFolder = pickFolder; _runner = runner ?? new OfficeWorkflowRunner();
        _localizer = localizer ?? StudioLocalization.Current; _storage = storage; _guard = guard; _jobs = jobs; _recovery = recovery;
        _confirmProviderWrite = confirmProviderWrite ?? (_ => Task.FromResult(false)); _openOutput = openOutput;
        Operations = [
            new(OfficeInvoiceWorkflowOperation.Inspect, T("Inspect", "Inspect")),
            new(OfficeInvoiceWorkflowOperation.Validate, T("Validate", "Validate")),
            new(OfficeInvoiceWorkflowOperation.Convert, T("Convert", "Convert XML")),
            new(OfficeInvoiceWorkflowOperation.RenderPresentationPdf, T("Render", "Presentation PDF")),
            new(OfficeInvoiceWorkflowOperation.RenderHybridPdf, T("Hybrid", "Hybrid PDF")),
            new(OfficeInvoiceWorkflowOperation.EditSource, T("Edit", "Edit source headers"))
        ];
        Targets = InvoiceXmlOptions.GetSupportedTargets().Select(options => new InvoiceTargetChoice(options,
            ReleaseLabel(options.Release) + " · " + options.Syntax + " · " + ProfileLabel(options.Profile))).ToArray();
        Releases = Targets.Select(target => target.Options.Release).Distinct().Select(release => new InvoiceReleaseChoice(release, ReleaseLabel(release))).ToArray();
        CodeDisplays = [new(InvoicePdfCodeDisplay.Code, T("Codes", "Codes")), new(InvoicePdfCodeDisplay.Description, T("Descriptions", "Descriptions")),
            new(InvoicePdfCodeDisplay.CodeAndDescription, T("CodesDescriptions", "Codes and descriptions"))];
        TableLayouts = [
            new(T("Table.Standard", "Standard"), [InvoicePdfLineColumn.Item, InvoicePdfLineColumn.Quantity, InvoicePdfLineColumn.NetPrice, InvoicePdfLineColumn.Vat, InvoicePdfLineColumn.NetAmount]),
            new(T("Table.Units", "Include units"), [InvoicePdfLineColumn.Item, InvoicePdfLineColumn.Quantity, InvoicePdfLineColumn.Unit, InvoicePdfLineColumn.NetPrice, InvoicePdfLineColumn.Vat, InvoicePdfLineColumn.NetAmount]),
            new(T("Table.Identifiers", "Include line identifiers"), [InvoicePdfLineColumn.LineIdentifier, InvoicePdfLineColumn.Item, InvoicePdfLineColumn.Quantity, InvoicePdfLineColumn.Unit, InvoicePdfLineColumn.NetPrice, InvoicePdfLineColumn.Vat, InvoicePdfLineColumn.NetAmount])
        ];
        _selectedOperation = Operations[0]; _selectedTarget = Targets[0]; _selectedRelease = Releases[0];
        _selectedUnitDisplay = CodeDisplays[1]; _selectedPaymentDisplay = CodeDisplays[1]; _selectedTableLayout = TableLayouts[1];
        _status = T("Ready", "Choose an invoice XML file and an operation.");
    }

    public IReadOnlyList<InvoiceOperationChoice> Operations { get; }
    public IReadOnlyList<InvoiceTargetChoice> Targets { get; }
    public IEnumerable<InvoiceTargetChoice> AvailableTargets => IsRendering ? Targets.Where(target => target.Options.Syntax == InvoiceSyntax.Cii) : Targets;
    public IReadOnlyList<InvoiceReleaseChoice> Releases { get; }
    public IReadOnlyList<InvoiceCodeDisplayChoice> CodeDisplays { get; }
    public IReadOnlyList<InvoiceTableLayoutChoice> TableLayouts { get; }
    public ObservableCollection<string> Diagnostics { get; } = new();
    public string Title => T("Title", "Invoices");
    public string Description => T("Description", "Inspect electronic invoice data, validate exact XML, create another format, or edit selected source headers.");
    public string EditingHint => T("EditingHint", "Blank fields retain their current values. Replacements require an existing unique field. Signed XML is protected; related references change only when supplied separately.");
    public bool IsStandardsAvailable => StudioDistributionPolicy.ExternalToolsAllowed;
    public string StandardsHint => IsStandardsAvailable
        ? T("StandardsHint", "Standards checks require the pinned rule files and a Java/Saxon runtime. Requested checks must pass before output is created.")
        : T("StandardsUnavailable", "External Java/Saxon standards checks are unavailable in the Mac App Store edition. XML inspection, model checks, conversion and PDF creation remain available.");
    public string InputFileName => string.IsNullOrWhiteSpace(InputPath) ? T("NoInput", "No invoice selected") : _storage?.Describe(InputPath).Name ?? OfficeStorageIdentity.GetFileName(InputPath);
    public bool IsRendering => SelectedOperation.Value is OfficeInvoiceWorkflowOperation.RenderPresentationPdf or OfficeInvoiceWorkflowOperation.RenderHybridPdf;
    public bool IsEditing => SelectedOperation.Value == OfficeInvoiceWorkflowOperation.EditSource;
    public bool IsWriting => IsRendering || IsEditing || SelectedOperation.Value == OfficeInvoiceWorkflowOperation.Convert;
    public bool NeedsTarget => IsWriting && !IsEditing;
    public bool ShowsSourceRelease => !NeedsTarget;
    public bool CanEdit => !_disposed && !IsBusy;
    public bool CanCancel => IsBusy;
    public bool HasOutput => !string.IsNullOrWhiteSpace(OutputPath);
    public bool CanOpenOutput => CanEdit && HasOutput && _openOutput != null;
    public bool CanRun => CanEdit && !string.IsNullOrWhiteSpace(InputPath) &&
        (!IsWriting || !string.IsNullOrWhiteSpace(OutputFolder) || OfficeStorageIdentity.GetLocalPath(InputPath) != null);

    [ObservableProperty] private InvoiceOperationChoice _selectedOperation;
    [ObservableProperty] private InvoiceTargetChoice _selectedTarget;
    [ObservableProperty] private InvoiceReleaseChoice _selectedRelease;
    [ObservableProperty] private InvoiceCodeDisplayChoice _selectedUnitDisplay;
    [ObservableProperty] private InvoiceCodeDisplayChoice _selectedPaymentDisplay;
    [ObservableProperty] private InvoiceTableLayoutChoice _selectedTableLayout;
    [ObservableProperty] private string _inputPath = string.Empty;
    [ObservableProperty] private string _outputFolder = string.Empty;
    [ObservableProperty] private string? _outputPath;
    [ObservableProperty] private bool _isBusy;
    [ObservableProperty] private bool _requireStandards;
    [ObservableProperty] private bool _allowProfileLoss;
    [ObservableProperty] private string _ruleBundlePath = string.Empty;
    [ObservableProperty] private string _peppolRulesPath = string.Empty;
    [ObservableProperty] private string _facturXRulesPath = string.Empty;
    [ObservableProperty] private string _saxonJarPath = string.Empty;
    [ObservableProperty] private string _javaExecutable = "java";
    [ObservableProperty] private string _languages = "en-US";
    [ObservableProperty] private string _fontPath = string.Empty;
    [ObservableProperty] private bool _modernLayout = true;
    [ObservableProperty] private bool _compactDetails = true;
    [ObservableProperty] private bool _pageIdentity = true;
    [ObservableProperty] private string _editNumber = string.Empty;
    [ObservableProperty] private string _editIssueDate = string.Empty;
    [ObservableProperty] private string _editDueDate = string.Empty;
    [ObservableProperty] private string _editBuyerReference = string.Empty;
    [ObservableProperty] private string _editPaymentReference = string.Empty;
    [ObservableProperty] private string _status;
    [ObservableProperty] private string _sourceSummary = "—";
    [ObservableProperty] private string _modelSummary = "—";
    [ObservableProperty] private string _mappingSummary = "—";
    [ObservableProperty] private string _schemaSummary = "—";
    [ObservableProperty] private string _rulesSummary = "—";

    partial void OnSelectedOperationChanged(InvoiceOperationChoice value) {
        if (IsRendering && SelectedTarget.Options.Syntax != InvoiceSyntax.Cii) SelectedTarget = Targets.First(target => target.Options.Syntax == InvoiceSyntax.Cii);
        foreach (string property in new[] { nameof(IsRendering), nameof(IsEditing), nameof(IsWriting), nameof(NeedsTarget), nameof(ShowsSourceRelease), nameof(AvailableTargets) }) OnPropertyChanged(property);
        NotifyCommands();
    }
    partial void OnInputPathChanged(string value) { OnPropertyChanged(nameof(InputFileName)); NotifyCommands(); }
    partial void OnOutputFolderChanged(string value) => NotifyCommands();
    partial void OnOutputPathChanged(string? value) { OnPropertyChanged(nameof(HasOutput)); OnPropertyChanged(nameof(CanOpenOutput)); OpenOutputCommand.NotifyCanExecuteChanged(); }
    partial void OnIsBusyChanged(bool value) => NotifyCommands();
    private void NotifyCommands() {
        foreach (string property in new[] { nameof(CanEdit), nameof(CanRun), nameof(CanCancel), nameof(CanOpenOutput) }) OnPropertyChanged(property);
        RunCommand.NotifyCanExecuteChanged(); CancelCommand.NotifyCanExecuteChanged(); ChooseInputCommand.NotifyCanExecuteChanged();
        ChooseOutputFolderCommand.NotifyCanExecuteChanged(); OpenOutputCommand.NotifyCanExecuteChanged();
    }
    private string T(string key, string fallback) => _localizer.GetOrDefault("Invoice." + key, fallback);
    private string ReleaseLabel(InvoiceSpecificationRelease release) => release switch {
        InvoiceSpecificationRelease.En16931_1_3_16 => "EN 16931 · 1.3.16",
        InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2 => "Factur-X 1.09.2 / ZUGFeRD 2.5.2",
        InvoiceSpecificationRelease.XRechnung_3_0_2_2026_08_31 => "XRechnung 3.0.2 · 2026-08-31",
        InvoiceSpecificationRelease.PeppolBis_3_0_21 => "Peppol BIS Billing · 3.0.21",
        _ => release.ToString()
    };
    private string ProfileLabel(InvoiceProfile profile) => profile == InvoiceProfile.BasicWithoutLines ? "BASIC WL" : profile.ToString();
    /// <summary>Stops accepting input and requests cancellation of the active execution.</summary>
    public void Dispose() { if (_disposed) return; _disposed = true; _cancellation?.Cancel(); NotifyCommands(); }
}

/// <summary>Operation value and localized selection label.</summary>
public sealed record InvoiceOperationChoice(OfficeInvoiceWorkflowOperation Value, string Label);
/// <summary>Owned authoring contract and its display label.</summary>
public sealed record InvoiceTargetChoice(InvoiceXmlOptions Options, string Label);
/// <summary>Explicit standards release for operations that retain source XML.</summary>
public sealed record InvoiceReleaseChoice(InvoiceSpecificationRelease Value, string Label);
/// <summary>Code-label presentation choice.</summary>
public sealed record InvoiceCodeDisplayChoice(InvoicePdfCodeDisplay Value, string Label);
/// <summary>Named line-column selection passed to the invoice PDF owner.</summary>
public sealed record InvoiceTableLayoutChoice(string Label, IReadOnlyList<InvoicePdfLineColumn> Columns);
