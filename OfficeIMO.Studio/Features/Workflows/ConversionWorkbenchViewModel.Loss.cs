using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace OfficeIMO.Studio.Features.Workflows;

/// <summary>
/// After a run, outputs whose fidelity evidence reports approximated or omitted content are named in one notice with a
/// way to their details, so a completed status is never the only thing the user sees about a lossy conversion.
/// </summary>
public sealed partial class ConversionWorkbenchViewModel {
    private readonly List<ConversionJobViewModel> _lossJobs = new();

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasLossNotice))]
    private string? _lossNotice;

    public bool HasLossNotice => LossNotice is not null;

    private void ResetLossNotice() {
        _lossJobs.Clear();
        LossNotice = null;
    }

    // Jobs being run again leave the notice until their new result is known; lossy results from earlier runs stay.
    private void ForgetLoss(IEnumerable<ConversionJobViewModel> rerun) {
        foreach (ConversionJobViewModel job in rerun) _lossJobs.Remove(job);
        RefreshLossNotice();
    }

    private void ReportLoss(IEnumerable<ConversionJobViewModel> completed) {
        foreach (ConversionJobViewModel job in completed.Where(job => job.HasLoss && !_lossJobs.Contains(job))) _lossJobs.Add(job);
        RefreshLossNotice();
    }

    // Removing or re-running a job takes it out of the notice; the notice closes when none remain.
    private void RefreshLossNotice() {
        _lossJobs.RemoveAll(job => !Jobs.Contains(job) || !job.HasLoss);
        LossNotice = _lossJobs.Count switch {
            0 => null,
            1 => _localizer.FormatOrDefault("Conversion.Loss.One",
                "{0} was converted, but some content was approximated or left out. Review the details before relying on the output.",
                _lossJobs[0].FileName),
            _ => _localizer.FormatOrDefault("Conversion.Loss.Many",
                "{0} files were converted with approximated or missing content. Review the details before relying on the outputs.",
                _lossJobs.Count)
        };
    }

    [RelayCommand]
    private void ReviewLoss() {
        RefreshLossNotice();
        if (_lossJobs.Count > 0) SelectedJob = _lossJobs[0];
    }

    [RelayCommand]
    private void DismissLossNotice() => ResetLossNotice();
}
