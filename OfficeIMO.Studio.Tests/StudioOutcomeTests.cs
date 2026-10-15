using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Localization;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Tests;

/// <summary>Convert, OCR and Jobs describe a finished task with the same states.</summary>
public sealed class StudioOutcomeTests {
    [Theory]
    [InlineData(null, false, StudioOutcome.None)]
    [InlineData(OfficeWorkflowStatus.Completed, false, StudioOutcome.Succeeded)]
    [InlineData(OfficeWorkflowStatus.Completed, true, StudioOutcome.CompletedWithNotes)]
    [InlineData(OfficeWorkflowStatus.Cancelled, false, StudioOutcome.Cancelled)]
    [InlineData(OfficeWorkflowStatus.Unconfirmed, false, StudioOutcome.NeedsChecking)]
    [InlineData(OfficeWorkflowStatus.Failed, false, StudioOutcome.Failed)]
    public void Workflow_statuses_map_to_one_set_of_outcomes(OfficeWorkflowStatus? status, bool hasNotes, StudioOutcome expected) =>
        Assert.Equal(expected, StudioOutcomes.From(status, hasNotes));

    [Fact]
    public void Output_that_could_not_be_confirmed_needs_checking_rather_than_counting_as_a_failure() {
        var history = new StudioJobHistory(StudioLocalization.Current);
        StudioJobRecord unconfirmed = history.Start("Conversion", "input.docx", "output.pdf", () => { });
        unconfirmed.Unconfirmed("The runner stopped before confirming the output.");
        Assert.True(unconfirmed.NeedsChecking);
        Assert.False(unconfirmed.IsFailed);
        Assert.False(unconfirmed.IsSucceeded);

        StudioJobRecord failed = history.Start("Conversion", "input.docx", "output.pdf", () => { });
        failed.Complete(OfficeWorkflowStatus.Failed, null, "The input could not be read.");
        Assert.True(failed.IsFailed);
        Assert.False(failed.NeedsChecking);
    }

    [Fact]
    public void A_conversion_with_warnings_or_loss_keeps_its_notes_in_job_history() {
        var history = new StudioJobHistory(StudioLocalization.Current);
        StudioJobRecord noted = history.Start("Conversion", "input.docx", "output.pdf", () => { });
        noted.Complete(OfficeWorkflowStatus.Completed, "output.pdf", "Converted with approximated content.", hasNotes: true);
        Assert.Equal(StudioOutcome.CompletedWithNotes, noted.SharedOutcome);
        Assert.True(noted.HasNotes);
        Assert.False(noted.IsSucceeded);
        Assert.True(noted.HasOutput);

        // Notes describe a completed result only.
        StudioJobRecord failed = history.Start("Conversion", "input.docx", "output.pdf", () => { });
        failed.Complete(OfficeWorkflowStatus.Failed, null, "The input could not be read.", hasNotes: true);
        Assert.Equal(StudioOutcome.Failed, failed.SharedOutcome);
    }
}
