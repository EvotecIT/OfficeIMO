using System;
using System.Collections.Generic;
namespace OfficeIMO.Archives;

/// <summary>
/// Describes a warning produced while traversing ZIP entries.
/// </summary>
public sealed class ArchiveTraversalWarning {
    /// <summary>
    /// Entry path related to the warning, when available.
    /// </summary>
    public string EntryPath { get; set; } = string.Empty;

    /// <summary>
    /// Warning message.
    /// </summary>
    public string Warning { get; set; } = string.Empty;
}
