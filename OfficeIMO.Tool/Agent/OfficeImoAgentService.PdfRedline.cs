using OfficeIMO.Diff;
using OfficeIMO.Pdf;
using OfficeIMO.Workflows;

namespace OfficeIMO.Tool.Agent;

internal sealed partial class OfficeImoAgentService {
    /// <summary>Text redline between two local PDF paths (CLI).</summary>
    internal Task<AgentPdfWorkflowResult> RedlinePdfAsync(string oldPath, string newPath, string outputPath, string? changesPath, string? reportPath,
        PdfWorkflowSettings settings, string? newPasswordEnvironmentVariable = null, bool ignoreCase = false, string algorithm = "myers",
        int maxOutputCharacters = 4000, CancellationToken cancellationToken = default) =>
        RedlineResolvedPdfAsync(ResolvePdfInput(oldPath), ResolvePdfInput(newPath), outputPath, changesPath, reportPath, settings,
            newPasswordEnvironmentVariable, ignoreCase, algorithm, maxOutputCharacters, cancellationToken);

    /// <summary>Text redline between two PDF sources registered by inspect or search (MCP).</summary>
    internal Task<AgentPdfWorkflowResult> RedlineRegisteredPdfAsync(string sourceId, string comparisonSourceId, string outputPath, string? changesPath,
        string? reportPath, PdfWorkflowSettings settings, string? newPasswordEnvironmentVariable = null, bool ignoreCase = false, string algorithm = "myers",
        int maxOutputCharacters = 4000, CancellationToken cancellationToken = default) =>
        RedlineResolvedPdfAsync(ResolveRegisteredPdfInput(sourceId, cancellationToken), ResolveRegisteredPdfInput(comparisonSourceId, cancellationToken),
            outputPath, changesPath, reportPath, settings, newPasswordEnvironmentVariable, ignoreCase, algorithm, maxOutputCharacters, cancellationToken);

    private async Task<AgentPdfWorkflowResult> RedlineResolvedPdfAsync(string oldInput, string newInput, string outputPath, string? changesPath,
        string? reportPath, PdfWorkflowSettings settings, string? newPasswordEnvironmentVariable, bool ignoreCase, string algorithm,
        int maxOutputCharacters, CancellationToken cancellationToken) {
        settings.Validate();
        maxOutputCharacters = ValidateOutputBudget(maxOutputCharacters);
        if (settings.Pages is not null) throw new AgentUsageException("A redline compares whole documents; page selection is not available.");
        OfficeDiffAlgorithm alignment = (algorithm ?? "myers").ToLowerInvariant() switch {
            "myers" => OfficeDiffAlgorithm.Myers,
            "patience" => OfficeDiffAlgorithm.Patience,
            "histogram" => OfficeDiffAlgorithm.Histogram,
            _ => throw new AgentUsageException("algorithm must be myers, patience, or histogram.")
        };
        var newSettings = new PdfWorkflowSettings { PasswordEnvironmentVariable = newPasswordEnvironmentVariable ?? settings.PasswordEnvironmentVariable };

        string[] sources = [oldInput, newInput];
        string destination = PreparePdfOutput(outputPath, sources, isDirectory: false, settings.Overwrite, "redline", maxOutputCharacters);
        string? changes = changesPath is null ? null : PreparePdfOutput(changesPath, sources, isDirectory: false, settings.Overwrite, "redline", maxOutputCharacters, ".json");
        string? report = reportPath is null ? null : PreparePdfOutput(reportPath, sources, isDirectory: false, settings.Overwrite, "redline", maxOutputCharacters, ".html");
        string?[] outputs = [destination, changes, report];
        if (outputs.Where(path => path is not null).Distinct(StringComparer.OrdinalIgnoreCase).Count() != outputs.Count(path => path is not null))
            throw new AgentUsageException("The annotated PDF, change list and report need separate output paths.");

        byte[] oldPdf = ReadBoundedInput(oldInput, settings.MaximumInputBytes);
        byte[] newPdf = ReadBoundedInput(newInput, settings.MaximumInputBytes);
        string? oldPassword = PdfPassword(settings);
        string? newPassword = PdfPassword(newSettings);
        var options = new PdfTextRedlineOptions {
            Algorithm = alignment,
            IgnoreCase = ignoreCase,
            OldName = Path.GetFileName(oldInput),
            NewName = Path.GetFileName(newInput),
            OldReadOptions = oldPassword is null ? null : new PdfLoadOptions { Password = oldPassword },
            NewReadOptions = newPassword is null ? null : new PdfLoadOptions { Password = newPassword }
        };

        PdfTextRedlineResult redline;
        try {
            redline = await Task.Run(() => PdfTextRedline.Compare(oldPdf, newPdf, options, cancellationToken), cancellationToken).ConfigureAwait(false);
        } catch (ArgumentException exception) when (exception.ParamName is "oldPdf" or "newPdf") {
            // The engine message names the side and the reason (for example a missing text layer); it carries no document text.
            throw new AgentUsageException(exception.Message.Split('\n')[0].Replace(" (Parameter '" + exception.ParamName + "')", string.Empty, StringComparison.Ordinal), exception);
        }

        // Every artifact is produced and the combined size checked before anything is written, so a refused run leaves no partial output.
        var pending = new List<(string Path, byte[] Content)> { (destination, redline.Pdf) };
        if (changes is not null) pending.Add((changes, System.Text.Encoding.UTF8.GetBytes(redline.ToJson())));
        if (report is not null) pending.Add((report, System.Text.Encoding.UTF8.GetBytes(redline.ToHtml(cancellationToken: cancellationToken))));
        if (pending.Sum(static item => item.Content.LongLength) > settings.MaximumOutputBytes) throw new AgentUsageException("The redline output exceeds maximumOutputBytes.");

        var artifacts = new List<AgentPdfArtifact>(pending.Count);
        foreach ((string path, byte[] content) in pending) {
            WriteRedlineOutput(path, content, settings.Overwrite);
            artifacts.Add(new AgentPdfArtifact { Path = path, SizeBytes = content.LongLength });
        }

        OfficeChangeSet set = redline.Changes;
        string summary = set.Changes.Count + " change(s): " + set.Count(OfficeChangeKind.Inserted) + " inserted, " +
            set.Count(OfficeChangeKind.Deleted) + " deleted, " + set.Count(OfficeChangeKind.Replaced) + " replaced, " +
            set.Count(OfficeChangeKind.Moved) + " moved; " + redline.AnnotationCount + " annotation(s) added." +
            (set.IsComplete ? string.Empty : " The alignment stopped at its work limit, so some unchanged text may be listed as changed.") +
            (redline.AnnotationLimitReached ? " Some changes are listed but not annotated because the annotation limit was reached." : string.Empty);
        return PdfResult("redline", OfficeWorkflowStatus.Completed, OfficeWorkflowFailureKind.None, destination, redline.Pdf.LongLength,
            artifacts, Array.Empty<OfficeWorkflowDiagnostic>(), maxOutputCharacters, summary: summary);
    }

    // Reads through one stream and stops at the limit, so a file that grows after it was opened cannot exceed it.
    private static byte[] ReadBoundedInput(string path, long maximumBytes) {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        if (stream.Length > maximumBytes) throw new AgentUsageException("An input PDF exceeds maximumInputBytes.");
        using var buffer = new MemoryStream((int)Math.Min(stream.Length, int.MaxValue));
        byte[] chunk = new byte[81920];
        int read;
        while ((read = stream.Read(chunk, 0, chunk.Length)) > 0) {
            if (buffer.Length + read > maximumBytes) throw new AgentUsageException("An input PDF exceeds maximumInputBytes.");
            buffer.Write(chunk, 0, read);
        }

        return buffer.ToArray();
    }

    private static void WriteRedlineOutput(string destination, byte[] content, bool overwrite) {
        string directory = Path.GetDirectoryName(destination)!;
        Directory.CreateDirectory(directory);
        string temporary = Path.Combine(directory, "." + Path.GetFileName(destination) + "." + Guid.NewGuid().ToString("N") + ".tmp");
        try {
            File.WriteAllBytes(temporary, content);
            File.Move(temporary, destination, overwrite);
        } finally {
            if (File.Exists(temporary)) File.Delete(temporary);
        }
    }
}
