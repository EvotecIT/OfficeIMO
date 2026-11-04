using System;
using System.Linq;
using OfficeIMO.Reader;
using PSWriteOffice.Models.Reader;
using PSWriteOffice.Services.Reader;

namespace PSWriteOffice.Cmdlets.Reader;

public sealed partial class ReadOfficeDocumentCommand {
    private void WriteProjection(OfficeDocumentReadResult document) {
        CancelToken.ThrowIfCancellationRequested();
        if (ReaderCommandUtilities.HasSourceLimit(document)) WriteWarning("Document processing reached a source limit and may be partial: " + document.Source?.Path);
        switch (As) {
            case OfficeReadView.Document: WriteObject(document); break;
            case OfficeReadView.Text:
                if (ByPage) WriteObject(document.GetPageText(CancelToken), true);
                else WriteObject(document.GetText(CancelToken));
                break;
            case OfficeReadView.Markdown:
                var pages = document.GetPageMarkdown(new OfficeDocumentPageMarkdownOptions { IncludePageMarkers = !NoPageMarkers });
                if (ByPage) WriteObject(pages, true);
                else WriteObject(pages.Count > 0 ? string.Join(Environment.NewLine + Environment.NewLine, pages.Select(page => page.Markdown)) : document.Markdown ?? document.GetText(CancelToken));
                break;
            case OfficeReadView.Json: WriteObject(document.ToJson(Indented)); break;
            case OfficeReadView.Chunks: WriteObject(document.Chunks, true); break;
            case OfficeReadView.Hierarchy: WriteObject(_reader.ChunkDocument(document, HierarchyOptions, CancelToken)); break;
            case OfficeReadView.Structure: WriteObject(document.ExtractStructured(StructureOptions, CancelToken)); break;
            case OfficeReadView.Tables:
                var tables = document.EnumerateTables().ToArray();
                if (Bundle) WriteObject(tables.ExportTables(Indented, CancelToken), true);
                else WriteObject(tables, true);
                break;
            case OfficeReadView.Visuals:
                var visuals = document.EnumerateVisuals().ToArray();
                if (Bundle) WriteObject(visuals.ExportVisuals(Indented, CancelToken), true);
                else WriteObject(visuals, true);
                break;
            case OfficeReadView.Assets: WriteObject(document.ExtractAssets(cancellationToken: CancelToken), true); break;
            case OfficeReadView.Ingest:
                var summary = new OfficeReadIngestSummary(NoChunks);
                summary.AddDocument(document);
                WriteObject(summary.Complete());
                break;
        }
    }
}
