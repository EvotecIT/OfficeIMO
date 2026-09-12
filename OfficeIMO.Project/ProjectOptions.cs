namespace OfficeIMO.Project;

/// <summary>Bounded XML input and common document lifecycle policy.</summary>
public sealed class ProjectLoadOptions : DocumentLoadOptions {
    /// <summary>Maximum input bytes, including buffering of non-seekable sources.</summary>
    public long MaxInputBytes { get; set; } = 64L * 1024 * 1024;
    /// <summary>Maximum compound directory entries for native input.</summary>
    public int MaxCompoundEntries { get; set; } = 4096;
    /// <summary>Maximum materialized compound streams for native input.</summary>
    public int MaxCompoundStreams { get; set; } = 2048;
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
    /// <summary>Maximum total variable field values materialized from native MPP/MPT tables.</summary>
    public int MaxNativeValues { get; set; } = 1_000_000;
    /// <summary>Maximum timephased intervals, without implicit expansion.</summary>
    public int MaxTimephasedValues { get; set; } = 1_000_000;
    /// <summary>Maximum retained diagnostics; truncation is itself diagnosed.</summary>
    public int MaxDiagnostics { get; set; } = 1_000;

    internal void ValidateLimits() {
        if (PackageSecurity != null)
            throw new NotSupportedException("PackageSecurity applies to packaged document formats. Use the Project XML limits for this codec; XML never resolves external entities or fetches external resources.");
        if (MaxInputBytes < 1 || MaxInputBytes > int.MaxValue || MaxCompoundEntries < 1 || MaxCompoundStreams < 1 || MaxCharacters < 1 || MaxDepth < 1 || MaxOutlineDepth < 1 ||
            MaxElements < 1 || MaxAttributes < 1 || MaxTasks < 1 || MaxEntities < 1 || MaxNativeValues < 1 ||
            MaxTimephasedValues < 1 || MaxDiagnostics < 1) throw new ArgumentOutOfRangeException(nameof(ProjectLoadOptions), "All limits must be positive and in-memory input must fit an Int32 byte array.");
        if (!Enum.IsDefined(typeof(DocumentAccessMode), AccessMode) || !Enum.IsDefined(typeof(DocumentPersistenceMode), PersistenceMode))
            throw new ArgumentOutOfRangeException(nameof(ProjectLoadOptions), "Unknown lifecycle policy.");
    }
}

/// <summary>A qualified Project serialization format.</summary>
public enum ProjectFileFormat {
    /// <summary>Use the destination extension, or retain the source format for stream saves. New streams use XML.</summary>
    Automatic,
    /// <summary>Microsoft Project XML (MSPDI).</summary>
    Xml,
    /// <summary>Modern MPP14 project document.</summary>
    Mpp14,
    /// <summary>Modern MPT14 document template; Global.mpt is a separate unsupported application store.</summary>
    Mpt14,
    /// <summary>Project 2007 MPP12 project document.</summary>
    Mpp12,
    /// <summary>Project 2007 MPT12 document template.</summary>
    Mpt12,
    /// <summary>Project 2000–2003 MPP9 project document.</summary>
    Mpp9,
    /// <summary>Project 2000–2003 MPT9 document template.</summary>
    Mpt9,
    /// <summary>Project 98 MPP8 project document.</summary>
    Mpp8,
    /// <summary>Project 98 MPT8 document template.</summary>
    Mpt8,
    /// <summary>Microsoft Project Exchange 4.0 text document.</summary>
    Mpx4
}

/// <summary>Deterministic legacy character mappings declared by MPX.</summary>
public enum ProjectMpxEncoding {
    /// <summary>ANSI interpreted as Windows code page 1252.</summary>
    Windows1252 = 1252,
    /// <summary>DOS code page 437.</summary>
    Dos437 = 437,
    /// <summary>DOS code page 850.</summary>
    Dos850 = 850,
    /// <summary>Macintosh Roman code page 10000.</summary>
    MacintoshRoman = 10000
}

/// <summary>Serialization format, fidelity, output bounds, and file conflict policy.</summary>
public sealed class ProjectSaveOptions {
    /// <summary>Target format. Explicit values must agree with a path's extension.</summary>
    public ProjectFileFormat Format { get; set; } = ProjectFileFormat.Automatic;
    /// <summary>MPX character mapping. Null retains an MPX source declaration; new MPX output uses Windows 1252. Unrepresentable characters are rejected.</summary>
    public ProjectMpxEncoding? MpxEncoding { get; set; }
    /// <summary>MPX field separator. Null retains the source separator or uses a comma for new output. Rewrites reject +, -, ., % and ? because these conflict with dependency syntax.</summary>
    public char? MpxSeparator { get; set; }
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
        if (MpxEncoding.HasValue && !Enum.IsDefined(typeof(ProjectMpxEncoding), MpxEncoding.Value)) throw new ArgumentOutOfRangeException(nameof(MpxEncoding));
        if (MpxSeparator.HasValue && (MpxSeparator.Value < 33 || MpxSeparator.Value > 126 || char.IsLetterOrDigit(MpxSeparator.Value) || MpxSeparator.Value == '"'))
            throw new ArgumentOutOfRangeException(nameof(MpxSeparator));
        if (!Enum.IsDefined(typeof(ProjectFileFormat), Format) || !Enum.IsDefined(typeof(OfficeConversionLossPolicy), LossPolicy) ||
            !Enum.IsDefined(typeof(OfficeConversionFileConflictPolicy), FileConflictPolicy) ||
            MaxOutputBytes < 1 || MaxOutputBytes > int.MaxValue) throw new ArgumentOutOfRangeException(nameof(ProjectSaveOptions));
    }
}
