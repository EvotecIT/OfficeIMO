using OfficeIMO.Core.Internal;
using OfficeIMO.Epub;
using OfficeIMO.Html;
using OfficeIMO.Provenance;

namespace OfficeIMO.Workflows;

public static partial class OfficeWorkflow {
    internal static bool IsManuscriptEpubRoute(string? routeId) =>
        routeId is "docx-epub" or "markdown-epub" or "html-epub";

    internal static async Task ValidateManuscriptArtifactAsync(Stream output, long maximumOutputBytes,
        EpubPublicationLoadOptions? retentionLimits, CancellationToken token) {
        EpubPublicationLoadOptions limits = retentionLimits ?? new EpubPublicationLoadOptions();
        var reading = new EpubLoadOptions {
            MaxPackageBytes = maximumOutputBytes,
            MaxArchiveEntries = limits.MaxEntries,
            MaxPackageMetadataBytes = Math.Max(limits.MaxMetadataBytes, limits.MaxEntryBytes),
            MaxChapters = limits.MaxEntries,
            MaxChapterBytes = limits.MaxEntryBytes,
            MaxTotalTextCharacters = limits.MaxExpandedBytes,
            MaxTotalUncompressedBytes = limits.MaxExpandedBytes,
            FallbackToHtmlScan = false
        };
        EpubDocument publication = await EpubDocument.LoadAsync(output, reading, token).ConfigureAwait(false);
        if (publication.Chapters.Count == 0 || !publication.ReadSummary.IsComplete ||
            publication.Diagnostics.Any(d => d.Severity == EpubDiagnosticSeverity.Error))
            throw new InvalidDataException("The converted EPUB could not be reopened with readable chapters.");
    }

    private sealed partial class PairRoute {
        private static async Task<OfficeConversionReport> ConvertManuscriptAsync(Stream input, Stream output,
            OfficeConversionOptions conversion, OfficeWorkflowRouteSettings settings, CancellationToken token) {
            EpubManuscriptOptions options = settings.EpubManuscript ?? new EpubManuscriptOptions();
            string sourceName = conversion.SourceName!;
            HtmlConversionDocumentOptions? htmlOptions = conversion.HtmlLoad?.Clone();
            if (conversion.ResourceDirectory != null) {
                string sourcePath = conversion.ResourceSourcePath ?? Path.Combine(conversion.ResourceDirectory, Path.GetFileName(sourceName));
                bool defaultPolicy = htmlOptions == null;
                htmlOptions ??= new HtmlConversionDocumentOptions();
                htmlOptions.BaseUri ??= new Uri(Path.GetFullPath(sourcePath));
                if (defaultPolicy) htmlOptions.ResourceUrlPolicy = OfficeHtmlResourceResolver.CreateResourcePolicy();
                if (conversion.ResourceSnapshots != null || options.ResourceResolver == null)
                    options.ResourceResolver = OfficeHtmlResourceResolver.CreateResolver(sourcePath, options.MaxResourceBytes, conversion.ResourceSnapshots);
            }
            EpubManuscriptResult imported = await BookManuscriptImporter.ImportStreamAsync(input, Path.GetExtension(sourceName), options,
                htmlOptions, Path.GetFileNameWithoutExtension(sourceName), forWrite: true, token,
                maximumXmlCharactersInPart: conversion.MaximumXmlCharactersInPart).ConfigureAwait(false);
            if (!imported.Succeeded || conversion.FailOnLoss) imported.Report.ThrowIfLoss();

            EpubWriteReport write;
            try {
                var save = new EpubSaveOptions {
                    MaxOutputBytes = conversion.MaximumOutputBytes,
                    MaxExpandedBytes = options.RetentionLimits.MaxExpandedBytes,
                    MaxEntries = options.RetentionLimits.MaxEntries
                };
                write = imported.Publication.WriteStaged(output, save, token);
            } catch (Exception error) when (error is (InvalidDataException or NotSupportedException or System.Xml.XmlException)
                && !OfficeProvenanceLimitException.IsOutput(error) && !OfficeOutputLimit.Is(error)) {
                var diagnostic = new OfficeDiagnostic("EPUB_IMPORT_PACKAGE_INVALID", OfficeDiagnosticSeverity.Error, error.Message,
                    sourceName, "OfficeIMO.Epub", OfficeConversionLossKind.Failure);
                var report = new OfficeConversionReport(sourceName, conversion.TargetName!,
                    diagnostics: imported.Report.Diagnostics.Concat(new[] { diagnostic }));
                throw new OfficeConversionException(error.Message, report, error);
            }
            return new OfficeConversionReport(sourceName, conversion.TargetName!, itemsWritten: 1,
                diagnostics: imported.Report.Diagnostics.Concat(write.Diagnostics));
        }
    }
}
