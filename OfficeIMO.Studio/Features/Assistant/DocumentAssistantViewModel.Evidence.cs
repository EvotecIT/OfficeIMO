using System.Diagnostics;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.AI;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Pdf;

namespace OfficeIMO.Studio.Features.Assistant;

internal sealed partial class DocumentAssistantViewModel {
    /// <summary>Prepares evidence locally. Only one immutable snapshot is retained while this assistant is active.</summary>
    [RelayCommand(CanExecute = nameof(CanPrepare))]
    private async Task PrepareEvidenceAsync() {
        if (!CanPrepare) return;
        if (_preparedDocument is not null && _source?.IsCurrent() == true) { Refresh(); return; }
        if (_source?.IsCurrent() != true) AllowRemoteProcessing = false;
        _preparedDocument = null; _readiness = null; EvidenceSummary = string.Empty;
        using var cancellation = new CancellationTokenSource(TimeSpan.FromMinutes(3));
        _operation = cancellation; IsBusy = true;
        int generation = _generation;
        try {
            Status = Text("Reading", "Reading the current document snapshot…");
            AssistantSource source = _capture(CurrentPageOnly);
            var timer = Stopwatch.StartNew();
            using var stream = new MemoryStream(source.Bytes, writable: false);
            OfficeAiDocument document = await OfficeAiDocument.ReadAsync(new OfficeDocumentReaderBuilder().AddPdfHandler(source.ReaderOptions).Build(),
                stream, source.Name, cancellationToken: cancellation.Token);
            if (_disposed || generation != _generation || cancellation.IsCancellationRequested || !source.IsCurrent()) return;
            if (_snapshotHash != document.SnapshotHash) { _history.Clear(); Messages.Clear(); LastAnswerText = string.Empty; }
            _source = source with { Bytes = [] };
            _snapshotHash = document.SnapshotHash;
            _preparedDocument = document;
            _readiness = OfficeAiEvidenceReadiness.Inspect(document, source.Page.HasValue ? [source.Page.Value] : null);
            EvidenceSummary = _localizer.FormatOrDefault("Assistant.EvidenceReady",
                "Pages: {0} · Text items: {1} · Characters: {2:N0}. Pages without extractable text: {3}. Prepared locally in {4:N0} ms.",
                _readiness.Pages.Count, _readiness.TextItems, _readiness.TextCharacters, _readiness.PagesWithoutText.Count, timer.Elapsed.TotalMilliseconds);
            if (_readiness.HasSourceDiagnostics) EvidenceSummary += " " + Text("EvidenceWarnings", "The document reader reported limitations. Review source coverage.");
            Status = _readiness.HasText ? Text("EvidenceAvailable", "Evidence is ready. It will be reused until the document or page scope changes. Nothing has been sent to a provider.")
                : Text("NeedsOcr", "No extractable text is available in this scope. Run OCR on scanned pages, review its output, then prepare evidence again.");
        } catch (OperationCanceledException) {
            if (!_disposed && generation == _generation) Status = Text("PreparationCancelled", "Evidence preparation cancelled. Nothing was sent to a provider.");
        } catch (AssistantSourceUnavailableException exception) {
            if (!_disposed && generation == _generation) Status = exception.Message;
        } catch (Exception) {
            if (!_disposed && generation == _generation) Status = Text("PreparationFailed", "This document could not be prepared within the reader limits. Check its protection, page count and text size, or use a smaller PDF.");
        } finally {
            if (ReferenceEquals(_operation, cancellation)) _operation = null;
            IsBusy = false;
        }
    }

    [RelayCommand(CanExecute = nameof(CanOpenOcr))]
    private async Task OpenOcrAsync() {
        if (!CanOpenOcr || _source?.IsCurrent() != true) return;
        using var cancellation = new CancellationTokenSource(TimeSpan.FromMinutes(3));
        _operation = cancellation; IsBusy = true;
        try { await OpenOcr!(cancellation.Token); }
        catch (OperationCanceledException) { Status = Text("PreparationCancelled", "Evidence preparation cancelled. Nothing was sent to a provider."); }
        catch (AssistantSourceUnavailableException exception) { Status = exception.Message; }
        catch (Exception) { Status = Text("OcrUnavailable", "OCR could not be opened. Save a local copy and open it in the OCR workspace."); }
        finally { if (ReferenceEquals(_operation, cancellation)) _operation = null; IsBusy = false; }
    }

    [RelayCommand(CanExecute = nameof(CanReviewAnswer))]
    private async Task CopyLastAnswerAsync() {
        if (!CanReviewAnswer || CopyAnswer is null) return;
        try {
            await CopyAnswer(LastAnswerText);
            Status = Text("AnswerCopied", "Answer, source quotes and limitations copied to the clipboard.");
        } catch (Exception) { Status = Text("CopyFailed", "The clipboard is unavailable. Export the answer instead."); }
    }

    [RelayCommand(CanExecute = nameof(CanReviewAnswer))]
    private async Task ExportLastAnswerAsync(CancellationToken token) {
        if (!CanReviewAnswer || ExportAnswer is null) return;
        using var cancellation = CancellationTokenSource.CreateLinkedTokenSource(token);
        _operation = cancellation; IsBusy = true;
        int generation = _generation;
        try {
            string text = LastAnswerText;
            bool saved = await ExportAnswer(text, () => !_disposed && !cancellation.IsCancellationRequested
                && generation == _generation && _source?.IsCurrent() == true && LastAnswerText == text, cancellation.Token);
            if (!_disposed && generation == _generation)
                Status = saved ? Text("AnswerExported", "Answer, source quotes and limitations exported.") : Text("ExportCancelled", "Answer export cancelled.");
        } catch (OperationCanceledException) { if (!_disposed && generation == _generation) Status = Text("ExportCancelled", "Answer export cancelled."); }
        catch (Exception) { if (!_disposed && generation == _generation) Status = Text("ExportFailed", "The answer could not be saved. Check the destination and try again."); }
        finally { if (ReferenceEquals(_operation, cancellation)) _operation = null; IsBusy = false; }
    }
}
