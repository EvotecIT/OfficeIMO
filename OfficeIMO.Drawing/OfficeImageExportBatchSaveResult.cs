using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace OfficeIMO.Drawing;

/// <summary>Paths, metadata, and aggregate diagnostics returned by a payload-releasing batch save.</summary>
public sealed class OfficeImageExportBatchSaveResult {
    private readonly ReadOnlyCollection<OfficeImageExportSavedFile> _files;

    internal OfficeImageExportBatchSaveResult(IEnumerable<OfficeImageExportSavedFile> files) {
        if (files == null) throw new ArgumentNullException(nameof(files));
        OfficeImageExportSavedFile[] snapshot = files.ToArray();
        _files = Array.AsReadOnly(snapshot);
        Report = new OfficeImageExportReport(
            snapshot.SelectMany(file => file.Diagnostics),
            snapshot.Length);
    }

    /// <summary>Saved files in export order. Encoded payloads are not retained.</summary>
    public IReadOnlyList<OfficeImageExportSavedFile> Files => _files;

    /// <summary>Aggregate diagnostics and fidelity status.</summary>
    public OfficeImageExportReport Report { get; }
}
