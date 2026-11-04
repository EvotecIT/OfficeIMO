using OfficeIMO.Drawing;
using System.Collections.Generic;
using System.Linq;
namespace OfficeIMO.OneNote.Html;

/// <summary>Visual-preserving HTML entry points backed by the native OneNote Drawing canvas.</summary>
public static partial class OneNoteVisualHtmlExtensions {
    private static string RenderVisual(string name, IEnumerable<OneNotePageReference> pages, OneNoteVisualHtmlExportOptions? options,
        out OfficeConversionReport report) {
        OneNoteVisualHtmlExportOptions operation = options?.Clone() ?? new OneNoteVisualHtmlExportOptions();
        var collected = new CollectingSink(operation.DiagnosticSink);
        operation.DiagnosticSink = collected;
        IReadOnlyList<OneNotePageReference> list = pages.ToList();
        string html = operation.Fragment
            ? OneNoteVisualHtmlRenderer.RenderFragment(list, operation)
            : OneNoteVisualHtmlRenderer.RenderDocument(name, list, operation);
        report = new OfficeConversionReport("OneNote", "Html", itemsWritten: list.Count,
            diagnostics: collected.Items.Select(item => new OfficeDiagnostic(item.Code,
                item.Severity == OfficeImageExportDiagnosticSeverity.Error ? OfficeDiagnosticSeverity.Error
                : item.Severity == OfficeImageExportDiagnosticSeverity.Warning ? OfficeDiagnosticSeverity.Warning
                : OfficeDiagnosticSeverity.Information, item.Message, item.Source, "OfficeIMO.OneNote.Html", item.LossKind)));
        return html;
    }

    private sealed class CollectingSink : ICollection<OfficeImageExportDiagnostic> {
        private readonly ICollection<OfficeImageExportDiagnostic>? _forward;
        internal CollectingSink(ICollection<OfficeImageExportDiagnostic>? forward) { _forward = forward; }
        internal List<OfficeImageExportDiagnostic> Items { get; } = new List<OfficeImageExportDiagnostic>();
        public int Count => Items.Count;
        public bool IsReadOnly => false;
        public void Add(OfficeImageExportDiagnostic item) { Items.Add(item); _forward?.Add(item); }
        public void Clear() => Items.Clear();
        public bool Contains(OfficeImageExportDiagnostic item) => Items.Contains(item);
        public void CopyTo(OfficeImageExportDiagnostic[] array, int arrayIndex) => Items.CopyTo(array, arrayIndex);
        public bool Remove(OfficeImageExportDiagnostic item) => Items.Remove(item);
        public IEnumerator<OfficeImageExportDiagnostic> GetEnumerator() => Items.GetEnumerator();
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => Items.GetEnumerator();
    }
}
