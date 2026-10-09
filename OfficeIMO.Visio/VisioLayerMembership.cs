using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace OfficeIMO.Visio;

/// <summary>Detached native membership cache, guarded by both layer names and ordinal positions.</summary>
internal sealed class VisioLayerMembership {
    private readonly HashSet<int> _indexes;
    private HashSet<string>? _names;
    internal XElement Cell { get; }
    internal bool HasBoundNames => _names != null;
    internal bool ProducerStateReplaced { get; private set; }

    internal VisioLayerMembership(XElement cell, IEnumerable<int> indexes) {
        Cell = new XElement(cell);
        _indexes = new HashSet<int>(indexes);
    }

    internal void BindNames(IEnumerable<string> names) => _names = new HashSet<string>(names, StringComparer.OrdinalIgnoreCase);

    internal bool IsCurrent(IEnumerable<string> names, IEnumerable<int> indexes) =>
        !ProducerStateReplaced && _names != null && _names.SetEquals(names) && _indexes.SetEquals(indexes);

    internal void ReplaceProducerState() => ProducerStateReplaced = true;

    internal VisioLayerMembership Clone() {
        var copy = new VisioLayerMembership(Cell, _indexes);
        copy.ProducerStateReplaced = ProducerStateReplaced;
        if (_names != null) copy.BindNames(_names);
        return copy;
    }
}
