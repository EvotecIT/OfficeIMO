using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Threading;
using Avalonia.VisualTree;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioObservedRunHistoryTests {
    [Fact]
    public async Task ObserverWindowClosesWithoutCancellingExternalRunAndRefreshUpdatesSelectedAttemptInPlace() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            services.Preferences.Update(current => current with { UiCulture = "en" });
            string source = Path.Combine(services.Paths.Root, "source");
            Directory.CreateDirectory(source);
            File.WriteAllText(Path.Combine(source, "one.txt"), "The owning run finishes independently");
            string jobPath = Path.Combine(services.Paths.Root, "job.json");
            var definition = new OfficeBatchOptions { InputDirectory = "source", OutputDirectory = "output" };
            (definition).Save(jobPath);
            string recoveryPath = CreateRecoveryFixture(services.Paths);
            byte[] recoveryBytes = File.ReadAllBytes(recoveryPath);
            string reportPath = services.DocumentHistory.ConversionRuns.CreateReportPath()!;
            using var cancellation = new CancellationTokenSource();
            int cancelRequests = 0;
            var owningRecord = services.Jobs.Start("Owning conversion", jobPath, "output", () => {
                cancelRequests++; cancellation.Cancel();
            }, batch: true);
            owningRecord.OwnRunReportPath(reportPath);
            var guard = new PausedPdfPublicationGuard();
            Task<OfficeBatchReport> execution = OfficeConverter.ConvertAsync(OfficeBatchOptions.Load(jobPath) with {
                ReportPath = reportPath, PublicationGuard = guard
            }, progress: null, cancellationToken: cancellation.Token);
            StudioApplicationServices? observingServices = null;
            MainWindow? window = null;
            try {
                await guard.Started.Task.WaitAsync(TimeSpan.FromSeconds(5));
                services.Jobs.RefreshRunReports();
                Assert.Same(owningRecord, Assert.Single(services.Jobs.Entries));
                Assert.True(owningRecord.IsActive); Assert.True(owningRecord.CanCancel);
                Assert.False(owningRecord.IsExternalRunning);
                observingServices = StudioApplicationServices.Create(services.Paths);
                var observed = Assert.Single(observingServices.Jobs.Entries, entry => entry.HasRunReport);
                var recovery = Assert.Single(observingServices.Jobs.Entries, entry => entry.HasRecovery);
                Assert.True(observed.IsExternalRunning);
                Assert.False(observed.IsActive); Assert.False(observed.CanCancel); Assert.False(observed.CanReviewJob);
                Assert.False(observed.CancelCommand.CanExecute(null));
                Assert.False(observingServices.Jobs.HasActive);
                Assert.Equal(0, observingServices.Jobs.ActiveCount);
                Assert.Equal(1, observingServices.Jobs.ExternalRunningCount);
                Assert.Equal(1, observingServices.Jobs.FinishedCount);
                Assert.Contains("elsewhere", observingServices.Jobs.Summary);
                Assert.Contains("1 finished", observingServices.Jobs.Summary);

                window = new MainWindow(observingServices) { Width = 960, Height = 620 };
                window.Show();
                window.ViewModel.ShowJobsCommand.Execute(null);
                Layout(window);
                var jobsView = window.GetVisualDescendants().OfType<StudioJobsView>().Single();
                var refresh = jobsView.GetVisualDescendants().OfType<Button>()
                    .Single(button => ReferenceEquals(button.Command, window.ViewModel.Jobs.RefreshHistoryCommand));
                Assert.True(refresh.IsEffectivelyVisible); Assert.True(refresh.IsEffectivelyEnabled);
                Point refreshLocation = refresh.TranslatePoint(default, window)!.Value;
                Assert.InRange(refreshLocation.X, 0, window.Bounds.Width - refresh.Bounds.Width);
                Assert.InRange(refreshLocation.Y, 0, window.Bounds.Height - refresh.Bounds.Height);
                refresh.Command!.Execute(null);
                Assert.Same(observed, Assert.Single(observingServices.Jobs.Entries, entry => entry.HasRunReport));
                Capture(window, "observed-jobs-running-960.png");
                Assert.False(window.TabHost.HasBusyDocuments);
                Assert.False(window.ViewModel.Jobs.ReviewJobCommand.CanExecute(observed));
                window.Close();
                await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
                Assert.False(window.IsVisible);
                Assert.Empty(window.OwnedWindows.OfType<ActiveOperationsDialog>());
                Assert.False(execution.IsCompleted);
                Assert.False(cancellation.IsCancellationRequested);
                Assert.Equal(0, cancelRequests);
                Assert.True(owningRecord.IsActive);
                Assert.True(OfficeBatchReport.IsActive(reportPath));

                using var jobs = new StudioJobsViewModel(observingServices.Jobs, (_, _) => Task.CompletedTask) {
                    ReviewSavedJob = (_, _, _) => Task.CompletedTask,
                    SelectedJob = observed
                };
                var local = observingServices.Jobs.Start("Local inspection", "local.txt", null, () => { });
                local.Complete(OfficeWorkflowStatus.Completed, null, "The local result stays unchanged");
                guard.Release.TrySetResult();
                var report = await execution.WaitAsync(TimeSpan.FromSeconds(10));
                owningRecord.CompleteBatch(OfficeWorkflowStatus.Completed, report.OutputDirectory,
                    "Owning conversion completed", [], hasVerifiedOutput: true, isDirectoryOutput: true);

                jobs.RefreshHistoryCommand.Execute(null);

                Assert.Same(observed, jobs.SelectedJob);
                Assert.Same(observed, Assert.Single(observingServices.Jobs.Entries, entry => entry.HasRunReport));
                Assert.Equal(report.RunId, observed.RunReport!.RunId);
                Assert.Equal(OfficeBatchRunState.Completed, observed.RunReport.State);
                Assert.False(observed.IsExternalRunning); Assert.False(observed.IsActive);
                Assert.True(observed.IsSucceeded); Assert.True(observed.CanReviewJob);
                Assert.True(jobs.ReviewJobCommand.CanExecute(observed));
                Assert.False(observingServices.Jobs.HasActive);
                Assert.Equal(0, observingServices.Jobs.ExternalRunningCount);
                Assert.Equal(3, observingServices.Jobs.FinishedCount);
                Assert.Same(local, Assert.Single(observingServices.Jobs.Entries, entry => entry.Title == "Local inspection"));
                Assert.Equal("The local result stays unchanged", local.Summary);
                Assert.Same(recovery, Assert.Single(observingServices.Jobs.Entries, entry => entry.HasRecovery));
                Assert.Equal(recoveryBytes, File.ReadAllBytes(recoveryPath));
                jobs.RefreshHistoryCommand.Execute(null);
                Assert.Equal(3, observingServices.Jobs.Entries.Count);
                Assert.Same(observed, jobs.SelectedJob);
                Assert.Equal(0, cancelRequests);
            } finally {
                guard.Release.TrySetResult();
                await execution.WaitAsync(TimeSpan.FromSeconds(10));
                if (owningRecord.IsActive) owningRecord.Complete(OfficeWorkflowStatus.Completed, null, "Test run ended");
                observingServices?.Jobs.RefreshRunReports();
                if (window is not null) {
                    foreach (Window dialog in window.OwnedWindows.ToArray()) dialog.Close(false);
                    window.Close();
                }
            }
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task AbandonedExternalLeaseRefreshesToInterruptedAndRemovingItsReportPreservesLocalWorkAndRecovery() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(() => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(Path.Combine(services.Paths.Root, "source"));
            string jobPath = Path.Combine(services.Paths.Root, "job.json");
            (new OfficeBatchOptions {
                InputDirectory = "source", OutputDirectory = "output", ResumeFrom = "checkpoints/plan.json"
            }).Save(jobPath);
            string jobSha256 = OfficeBatchOptions.Load(jobPath).JobSha256!;
            string recoveryPath = CreateRecoveryFixture(services.Paths);
            byte[] recoveryBytes = File.ReadAllBytes(recoveryPath), jobBytes = File.ReadAllBytes(jobPath);
            string reportPath = services.DocumentHistory.ConversionRuns.CreateReportPath()!;
            var report = new OfficeBatchReport {
                RunId = Guid.NewGuid().ToString("N"), StartedUtc = DateTimeOffset.UtcNow,
                JobPath = jobPath, JobSha256 = jobSha256, OutputDirectory = Path.Combine(services.Paths.Root, "output"),
                ResumeFrom = Path.Combine(services.Paths.Root, "checkpoints", "plan.json"), State = OfficeBatchRunState.Running, Summary = new(0, 0, 0, 0, false)
            };
            File.WriteAllText(reportPath, OfficeBatchReportSerializer.Serialize(report));
            var lease = new FileStream(reportPath + ".lease", FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None);
            lease.Write(Encoding.ASCII.GetBytes(report.RunId)); lease.Flush(flushToDisk: true);
            StudioJobRecord? local = null;
            try {
                var observingServices = StudioApplicationServices.Create(services.Paths);
                var history = observingServices.Jobs;
                var observed = Assert.Single(history.Entries, entry => entry.HasRunReport);
                var recovery = Assert.Single(history.Entries, entry => entry.HasRecovery);
                using var jobs = new StudioJobsViewModel(history, (_, _) => Task.CompletedTask) {
                    ReviewSavedJob = (_, _, _) => Task.CompletedTask,
                    SelectedJob = observed
                };
                Assert.False(history.HasActive);
                Assert.Equal(1, history.ExternalRunningCount);
                int cancelRequests = 0;
                local = history.Start("Local active conversion", "local.txt", null, () => cancelRequests++);
                Assert.True(local.IsActive); Assert.True(local.CanCancel);
                Assert.False(history.CanClear); Assert.False(jobs.ClearFinishedCommand.CanExecute(null));
                Assert.False(observed.CanReviewJob); Assert.False(observed.CanCancel);

                history.ClearFinished();

                Assert.Equal(3, history.Entries.Count);
                Assert.Contains(observed, history.Entries); Assert.True(File.Exists(reportPath));
                Assert.True(OfficeBatchReport.IsActive(reportPath));
                observingServices.DocumentHistory.SetRememberHistory(false);
                var laterReport = report with {
                    RunId = Guid.NewGuid().ToString("N"), State = OfficeBatchRunState.Completed,
                    FinishedUtc = report.StartedUtc.AddSeconds(1),
                    Summary = new(0, 0, 0, 0, false), Succeeded = true
                };
                string laterReportPath = Path.Combine(services.Paths.ConversionRunsRoot, Guid.NewGuid().ToString("N") + ".json");
                File.WriteAllText(laterReportPath, OfficeBatchReportSerializer.Serialize(laterReport));
                lease.Dispose();
                jobs.RefreshHistoryCommand.Execute(null);
                Assert.Same(observed, jobs.SelectedJob);
                Assert.Same(observed, Assert.Single(history.Entries, entry => entry.HasRunReport));
                Assert.Equal(OfficeBatchRunState.Interrupted, observed.RunReport!.State);
                Assert.False(observed.IsExternalRunning); Assert.False(observed.IsActive);
                Assert.True(observed.CanReviewJob); Assert.True(jobs.ReviewJobCommand.CanExecute(observed));
                Assert.False(observed.CanCancel); Assert.Equal(0, history.ExternalRunningCount);
                Assert.Equal(1, history.ActiveCount); Assert.Equal(2, history.FinishedCount);
                Assert.True(local.IsActive); Assert.Equal(0, cancelRequests);
                Assert.DoesNotContain(history.Entries, entry => entry.RunReport?.RunId == laterReport.RunId);

                Assert.True(observingServices.DocumentHistory.ConversionRuns.RemoveFinished(reportPath));
                jobs.RefreshHistoryCommand.Execute(null);

                Assert.Null(jobs.SelectedJob);
                Assert.DoesNotContain(observed, history.Entries);
                Assert.Equal(2, history.Entries.Count);
                Assert.Same(local, Assert.Single(history.Entries, entry => entry.IsActive));
                Assert.Same(recovery, Assert.Single(history.Entries, entry => entry.HasRecovery));
                Assert.Equal(1, history.ActiveCount); Assert.Equal(0, history.ExternalRunningCount); Assert.Equal(1, history.FinishedCount);
                Assert.Equal(0, cancelRequests);
                Assert.Equal(recoveryBytes, File.ReadAllBytes(recoveryPath));
                Assert.Equal(jobBytes, File.ReadAllBytes(jobPath));
                Assert.True(File.Exists(laterReportPath));
            } finally {
                lease.Dispose();
                if (local?.IsActive == true) local.Complete(OfficeWorkflowStatus.Cancelled, null, "Test run ended");
            }
            return true;
        }, CancellationToken.None);
    }

    private static string CreateRecoveryFixture(StudioDataPaths paths) {
        string id = Guid.NewGuid().ToString("N"), directory = Path.Combine(paths.WorkflowRecoveryRoot, "output-" + id);
        Directory.CreateDirectory(directory);
        string artifact = Path.Combine(directory, "output.pdf");
        PdfDocument.Create(document => document.Page(page => page.Size(global::OfficeIMO.OfficeLength.Points(200), global::OfficeIMO.OfficeLength.Points(300)))).Save(artifact);
        byte[] bytes = File.ReadAllBytes(artifact);
        File.WriteAllText(Path.Combine(directory, "record.json"), JsonSerializer.Serialize(new {
            Version = 1, Id = id, Name = "recovered.pdf", Destination = Path.Combine(paths.Root, "recovered.pdf"),
            Length = bytes.LongLength, Sha256 = Convert.ToHexString(SHA256.HashData(bytes)), CreatedUtc = DateTimeOffset.UtcNow
        }));
        return artifact;
    }

    private static void Layout(Window window) {
        window.UpdateLayout(); Dispatcher.UIThread.RunJobs(); AvaloniaHeadlessPlatform.ForceRenderTimerTick(); window.UpdateLayout();
    }

    private static void Capture(Window window, string name) {
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrWhiteSpace(output)) return;
        Directory.CreateDirectory(output);
        Layout(window);
        using var frame = window.CaptureRenderedFrame(); Assert.NotNull(frame);
        frame.Save(Path.Combine(output, name), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }


}
