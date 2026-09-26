using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Features.Shell;

/// <summary>Standards readiness: checks the open document against archival, accessibility and print profiles.</summary>
public sealed partial class MainWindowViewModel {
    private long _complianceRevision = -1;

    public IReadOnlyList<ComplianceProfileChoice> ComplianceProfiles { get; } = [
        new(PdfComplianceProfile.PdfA2B, "PDF/A-2b"),
        new(PdfComplianceProfile.PdfA2U, "PDF/A-2u"),
        new(PdfComplianceProfile.PdfA3B, "PDF/A-3b"),
        new(PdfComplianceProfile.PdfA4, "PDF/A-4"),
        new(PdfComplianceProfile.PdfUa1, "PDF/UA-1"),
        new(PdfComplianceProfile.PdfUa2, "PDF/UA-2"),
        new(PdfComplianceProfile.PdfX4, "PDF/X-4")
    ];

    [ObservableProperty]
    private ComplianceProfileChoice? _selectedComplianceProfile;

    [ObservableProperty]
    private string? _complianceSummary;

    [ObservableProperty]
    private bool _isComplianceReady;

    public ObservableCollection<ComplianceRequirementViewModel> ComplianceIssues { get; } = [];

    public bool HasComplianceResult => ComplianceSummary is not null;

    partial void OnComplianceSummaryChanged(string? value) => OnPropertyChanged(nameof(HasComplianceResult));

    partial void OnSelectedComplianceProfileChanged(ComplianceProfileChoice? value) => ClearComplianceResult();

    private void ClearComplianceResult() {
        ComplianceIssues.Clear();
        ComplianceSummary = null;
        IsComplianceReady = false;
        _complianceRevision = -1;
    }

    // A result describes one revision; any edit makes it stale.
    private void InvalidateComplianceResult() {
        if (_workspace is null || _workspace.Revision != _complianceRevision) ClearComplianceResult();
    }

    private bool CanCheckCompliance() => _workspace is not null && !IsWorkspaceBusy;

    [RelayCommand(CanExecute = nameof(CanCheckCompliance))]
    private async Task CheckComplianceAsync(CancellationToken cancellationToken) {
        if (_workspace is not { } workspace) return;
        ComplianceProfileChoice profile = SelectedComplianceProfile ??= ComplianceProfiles[0];
        PdfComplianceReadinessReport? report = null;
        bool succeeded = await RunStandaloneAsync(async token => report = await workspace.AssessComplianceAsync(profile.Profile, token).ConfigureAwait(true),
            cancellationToken, describeSuccess: () => report is null ? null : UiFormat("Compliance.Checked", profile.Label)).ConfigureAwait(true);
        if (!succeeded || report is null || !ReferenceEquals(workspace, _workspace) ||
            !ReferenceEquals(profile, SelectedComplianceProfile)) return;
        ComplianceIssues.Clear();
        foreach (PdfComplianceRequirement requirement in report.Requirements.Where(item => item.Status != PdfComplianceRequirementStatus.Satisfied))
            ComplianceIssues.Add(new ComplianceRequirementViewModel(requirement.DisplayName, requirement.Diagnostic,
                requirement.Status == PdfComplianceRequirementStatus.Unsupported));
        IsComplianceReady = report.IsReady;
        int met = report.Requirements.Count(item => item.Status == PdfComplianceRequirementStatus.Satisfied);
        ComplianceSummary = report.IsReady
            ? UiFormat("Compliance.Ready", profile.Label, met)
            : UiFormat("Compliance.NotReady", profile.Label, ComplianceIssues.Count, met);
        _complianceRevision = workspace.Revision;
    }
}

public sealed record ComplianceProfileChoice(PdfComplianceProfile Profile, string Label);

public sealed record ComplianceRequirementViewModel(string Title, string Detail, bool IsUnsupported);
