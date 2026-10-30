using System;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Services.PowerPoint;

/// <summary>Collects semantic blocks backed by live native bodies in one rollback-owned slide.</summary>
internal sealed class PowerPointContentScope : IDisposable {
    private readonly PowerPointLayoutMaterialization? _materialization;
    private readonly List<PowerPointLayoutNativeText> _bodies = new();
    private readonly List<PowerPointLayoutTable> _tables = new();
    private readonly List<PowerPointShape> _figures = new();
    private readonly Action? _rollbackNativeState;
    private bool _accepted;

    internal PowerPointContentScope(PowerPointPresentation document, PowerPointContentScope? parent = null) {
        Document = document;
        Slide = parent?.Slide ?? this;
        if (parent == null) _materialization = new PowerPointLayoutMaterialization(document);
        else {
            Action rollbackNotes = PowerPointLayoutMaterialization.CaptureNotesRollback(NativeSlide);
            Action rollbackSections = document.CaptureSectionRollback();
            _rollbackNativeState = () => { rollbackNotes(); rollbackSections(); };
        }
    }
    internal PowerPointPresentation Document { get; }
    internal PowerPointContentScope Slide { get; }
    internal PowerPointSlide NativeSlide => Slide._materialization!.Slide;
    internal List<PowerPointLayoutBlock> Blocks { get; } = new();

    internal static PowerPointContentScope? Current {
        get {
            var scope = OfficeDocumentContext.Find<PowerPointContentScope>();
            return scope != null && ReferenceEquals(scope.Document, OfficeDocumentContext.Find<PowerPointPresentation>()) ? scope : null;
        }
    }

    internal static void Add(PowerPointLayoutBlock block) {
        var scope = Current ?? throw new PSArgumentException("Invoke this content command inside PowerPointSlide.");
        scope.Blocks.Add(block);
    }

    internal PowerPointLayoutNativeText CreateTextBody() {
        var body = new PowerPointLayoutNativeText(Slide._materialization!.CreateTextBody());
        _bodies.Add(body);
        return body;
    }

    internal void DiscardTextBody(PowerPointLayoutNativeText body) {
        _bodies.Remove(body);
        body.Box.Remove();
    }

    internal void AddTable(PowerPointLayoutTable table) {
        _tables.Add(table);
        Blocks.Add(table);
    }

    internal void AddFigure(PowerPointLayoutBlock block, PowerPointShape figure) {
        _figures.Add(figure);
        Blocks.Add(block);
    }

    internal void AcceptColumn(PowerPointContentScope column, string? title) {
        // Validate before transferring ownership; a rejected column rolls back its own native frames and notes.
        var block = new PowerPointLayoutColumn(title, column.Blocks);
        Blocks.Add(block);
        _bodies.AddRange(column._bodies);
        _tables.AddRange(column._tables);
        _figures.AddRange(column._figures);
        column._accepted = true;
    }

    internal IReadOnlyList<PowerPointSlide> Complete(PowerPointContentPlanSlide plan, PowerPointCompositionOptions options,
        Action<IReadOnlyList<PowerPointSlide>>? finalize = null) {
        foreach (var body in _bodies) body.Freeze();
        foreach (var table in _tables) table.FreezeNativeRows();
        var result = _materialization!.Complete(plan, options, finalize);
        _accepted = true;
        return result;
    }

    public void Dispose() {
        if (_materialization != null) _materialization.Dispose();
        else if (!_accepted) {
            foreach (var body in _bodies) body.Box.Remove();
            foreach (var table in _tables) table.NativeBody!.Table.Remove();
            foreach (var figure in _figures) figure.Remove();
            _rollbackNativeState!();
        }
    }
}
