namespace OfficeIMO.Project;

/// <summary>Bounded XML input and common document lifecycle policy.</summary>
public sealed class ProjectLoadOptions : DocumentLoadOptions {
    /// <summary>Maximum input bytes, including buffering of non-seekable sources.</summary>
    public long MaxInputBytes { get; set; } = 64L * 1024 * 1024;
    /// <summary>Maximum XML characters before materialization.</summary>
    public long MaxCharacters { get; set; } = 64L * 1024 * 1024;
    /// <summary>Maximum XML nesting depth.</summary>
    public int MaxDepth { get; set; } = 128;
    /// <summary>Maximum task outline level, independent of XML element nesting.</summary>
    public int MaxOutlineDepth { get; set; } = 128;
    /// <summary>Maximum total XML elements.</summary>
    public int MaxElements { get; set; } = 2_000_000;
    /// <summary>Maximum total XML attributes.</summary>
    public int MaxAttributes { get; set; } = 2_000_000;
    /// <summary>Maximum tasks, including summary and null placeholder records.</summary>
    public int MaxTasks { get; set; } = 100_000;
    /// <summary>Maximum combined task/resource/calendar/assignment entities.</summary>
    public int MaxEntities { get; set; } = 300_000;
    /// <summary>Maximum timephased intervals, without implicit expansion.</summary>
    public int MaxTimephasedValues { get; set; } = 1_000_000;
    /// <summary>Maximum retained diagnostics; truncation is itself diagnosed.</summary>
    public int MaxDiagnostics { get; set; } = 1_000;

    internal void ValidateLimits() {
        if (PackageSecurity != null)
            throw new NotSupportedException("PackageSecurity applies to packaged document formats. Use the Project XML limits for this codec; XML never resolves external entities or fetches external resources.");
        if (MaxInputBytes < 1 || MaxInputBytes > int.MaxValue || MaxCharacters < 1 || MaxDepth < 1 || MaxOutlineDepth < 1 ||
            MaxElements < 1 || MaxAttributes < 1 || MaxTasks < 1 || MaxEntities < 1 ||
            MaxTimephasedValues < 1 || MaxDiagnostics < 1) throw new ArgumentOutOfRangeException(nameof(ProjectLoadOptions), "All limits must be positive and in-memory input must fit an Int32 byte array.");
        if (!Enum.IsDefined(typeof(DocumentAccessMode), AccessMode) || !Enum.IsDefined(typeof(DocumentPersistenceMode), PersistenceMode))
            throw new ArgumentOutOfRangeException(nameof(ProjectLoadOptions), "Unknown lifecycle policy.");
    }
}

/// <summary>XML serialization, fidelity, output bounds, and file conflict policy.</summary>
public sealed class ProjectSaveOptions {
    /// <summary>Blocks save when modeled edits may invalidate unmodeled source content.</summary>
    public OfficeConversionLossPolicy LossPolicy { get; set; } = OfficeConversionLossPolicy.Block;
    /// <summary>Existing-file behavior. Associated saves replace by default through the shared atomic writer.</summary>
    public OfficeConversionFileConflictPolicy FileConflictPolicy { get; set; } = OfficeConversionFileConflictPolicy.Replace;
    /// <summary>Returns original bytes when no edits were made.</summary>
    public bool PreserveUnchangedBytes { get; set; } = true;
    /// <summary>Indents newly serialized XML.</summary>
    public bool Indent { get; set; } = true;
    /// <summary>Maximum serialized output bytes.</summary>
    public long MaxOutputBytes { get; set; } = 128L * 1024 * 1024;
    internal void Validate() {
        if (!Enum.IsDefined(typeof(OfficeConversionLossPolicy), LossPolicy) ||
            !Enum.IsDefined(typeof(OfficeConversionFileConflictPolicy), FileConflictPolicy) ||
            MaxOutputBytes < 1 || MaxOutputBytes > int.MaxValue) throw new ArgumentOutOfRangeException(nameof(ProjectSaveOptions));
    }
}
