using System;

namespace OfficeIMO.Visio;

public partial class VisioDocument : IDisposable {
    /// <summary>Completes the document lifecycle without saving or closing caller-owned streams.</summary>
    /// <remarks>The document retains its package in memory and owns no open file handles.</remarks>
    public void Dispose() { }
}
