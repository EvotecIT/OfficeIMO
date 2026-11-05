using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO;
using OfficeIMO.Internal;
using OfficeIMO.Reader;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Reader;

namespace PSWriteOffice.Cmdlets.Conversion;

public sealed partial class ConvertOfficeDocumentCommand {
    private async Task ConvertOcrBatchAsync() {
        if (string.IsNullOrWhiteSpace(OutputDirectory)) throw new PSArgumentException("Specify OutputDirectory for this batch.");
        string outputDirectory = OfficePath.ResolveOutput(SessionState, OutputDirectory!, directory: true);
        string target = OfficeFormatSelection.Parse(To)!.Extension;
        var options = CaptureOptions(_inputs, target);
        string[] directories = _inputs.Where(Directory.Exists).ToArray();
        if (directories.Length > 1 || directories.Length == 1 && _inputs.Count != 1)
            throw new PSArgumentException("Select one directory or a set of files for an OCR batch.");
        string? sourceDirectory = directories.SingleOrDefault();
        if (sourceDirectory != null && (OfficePathIdentity.IsSameOrDescendant(outputDirectory, sourceDirectory) ||
            OfficePathIdentity.IsSameOrDescendant(sourceDirectory, outputDirectory)))
            throw new PSArgumentException("OCR input and output directory trees must not overlap.");
        OfficeBatchExecutor.EnsureNoLinks(outputDirectory);
        int maximum = MaximumFiles == 0 ? int.MaxValue - 1 : MaximumFiles;
        var folder = new ReaderFolderOptions {
            Recurse = Recurse.IsPresent, MaxFiles = maximum + 1,
            Extensions = Extension?.Select(NormalizeExtension).ToArray() ??
                OfficeIMO.Drawing.OfficeImageReader.KnownExtensions.Concat(new[] { ".pdf" }).ToArray()
        };
        var paths = ReaderCommandUtilities.Reader.EnumerateDocumentPaths(_inputs, folder, CancelToken)
            .Take(maximum + 1).Distinct(StringComparer.Ordinal).ToArray();
        if (paths.Length > maximum) throw new PSArgumentException("The input selection exceeds MaximumFiles.");
        var items = new List<(string Input, string Output)>();
        var claims = new HashSet<string>(StringComparer.Ordinal);
        foreach (string input in paths) {
            CancelToken.ThrowIfCancellationRequested();
            if (Directory.Exists(input)) throw new PSArgumentException("OCR requires original image or PDF files, not directory packages.");
            string name = sourceDirectory == null ? System.IO.Path.GetFileName(input) : OfficeRuntimeCompatibility.GetRelativePath(sourceDirectory, input);
            string output = System.IO.Path.Combine(outputDirectory, System.IO.Path.ChangeExtension(name, target));
            OfficeBatchExecutor.EnsureNoLinks(input);
            OfficeBatchExecutor.EnsureNoLinks(output);
            if (paths.Any(source => OfficePath.AreEquivalent(source, output)))
                throw new PSArgumentException("An OCR destination must not replace any selected source.");
            if (!claims.Add(OfficePathIdentity.GetPathIdentityKey(output)))
                throw new PSArgumentException("Multiple OCR inputs select the same destination.");
            items.Add((input, output));
        }
        if (items.Count == 0 && !AllowEmpty) throw new PSArgumentException("No files were selected for OCR conversion.");
        if (!ShouldProcess(outputDirectory, "Recognize and convert document batch")) return;
        long completed = 0, failed = 0, skipped = 0;
        try {
            await OfficeBatch.RunAsync(items, async (item, token) => {
                try {
                    if (OfficeBatchOutputSelection.ObserveSkip(item.Input, item.Output, NoClobber.IsPresent, SkipUpToDate.IsPresent) is { } observation) {
                        Interlocked.Increment(ref skipped);
                        return (object?)observation;
                    }
                    var result = await ConvertOcrPathAsync(item.Input, item.Output, options, token).ConfigureAwait(false);
                    Interlocked.Increment(ref completed);
                    OpenPublishedOutput(result);
                    return result;
                } catch (OperationCanceledException) when (token.IsCancellationRequested) { throw; }
                catch { Interlocked.Increment(ref failed); throw; }
            }, ThrottleLimit, CancelToken, this, "Recognize and convert documents", "conversion-failed");
        } finally {
            SetSummary(new OfficeBatchSummary(items.Count, completed, 0, failed, CancelToken.IsCancellationRequested, skipped));
        }
    }
}
