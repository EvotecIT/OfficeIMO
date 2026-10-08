namespace OfficeIMO.Access;

/// <summary>Access file family; a native generation is selected separately.</summary>
public enum AccessFileFormat {
    /// <summary>Jet database (.mdb).</summary>
    Mdb,
    /// <summary>ACE database (.accdb).</summary>
    Accdb
}

/// <summary>Physical generation declared by a database header. Detection does not imply codec support.</summary>
public enum AccessFormatProfile {
    /// <summary>Unrecognized generation.</summary>
    Unknown,
    /// <summary>Jet 3, normally Access 97.</summary>
    Jet3,
    /// <summary>Jet 4, Access 2000–2003.</summary>
    Jet4,
    /// <summary>ACE generation introduced with Access 2007.</summary>
    Ace12,
    /// <summary>ACE generation introduced with Access 2010.</summary>
    Ace14,
    /// <summary>ACE generation introduced with Access 2016.</summary>
    Ace16,
    /// <summary>ACE generation introduced with Access 2019.</summary>
    Ace17
}

/// <summary>Options for a new in-memory database model. Creation does not produce a native file.</summary>
public sealed class AccessCreateOptions : DocumentCreateOptions {
    /// <summary>Target file family. New ACCDB models select ACE 12; MDB models select Jet 4.</summary>
    public AccessFileFormat Format { get; set; } = AccessFileFormat.Accdb;
    /// <summary>Optional explicit generation. New models target Jet4 for MDB and Ace12 for ACCDB.</summary>
    public AccessFormatProfile? Profile { get; set; }
    internal void Validate() {
        if (!Enum.IsDefined(typeof(AccessFileFormat), Format)) throw new ArgumentOutOfRangeException(nameof(Format));
        if (Profile.HasValue && Profile.Value != (Format == AccessFileFormat.Mdb ? AccessFormatProfile.Jet4 : AccessFormatProfile.Ace12))
            throw new NotSupportedException("New Access models currently target Jet4 MDB or Ace12 ACCDB; the explicit profile must agree with that family.");
        if (!Enum.IsDefined(typeof(DocumentPersistenceMode), PersistenceMode)) throw new ArgumentOutOfRangeException(nameof(PersistenceMode));
        // Do not promise disposal persistence before a native writer is qualified.
        if (PersistenceMode != DocumentPersistenceMode.Explicit) throw new NotSupportedException("Access native persistence is not qualified; use Explicit persistence and AssessSave.");
    }
}

/// <summary>Bounded native inspection and document ownership options.</summary>
public sealed class AccessLoadOptions : DocumentLoadOptions {
    /// <summary>Maximum bytes snapshotted from input, including a non-seekable stream.</summary>
    public long MaxInputBytes { get; set; } = 64L * 1024 * 1024;
    /// <summary>Maximum physical pages accepted before any catalog decoding.</summary>
    public int MaxPages { get; set; } = 16_384;
    internal void Validate() {
        if (MaxInputBytes < 1 || MaxInputBytes > int.MaxValue) throw new ArgumentOutOfRangeException(nameof(MaxInputBytes));
        if (MaxPages < 1) throw new ArgumentOutOfRangeException(nameof(MaxPages));
        if (!Enum.IsDefined(typeof(DocumentAccessMode), AccessMode)) throw new ArgumentOutOfRangeException(nameof(AccessMode));
        if (!Enum.IsDefined(typeof(DocumentPersistenceMode), PersistenceMode)) throw new ArgumentOutOfRangeException(nameof(PersistenceMode));
        if (PackageSecurity != null) throw new NotSupportedException("Access databases are paged database files. PackageSecurity is not an Access protection policy.");
        OfficeIMO.Core.Internal.OfficeDocumentLifecycle.Validate(AccessMode, PersistenceMode, "Access document");
        if (PersistenceMode != DocumentPersistenceMode.Explicit) throw new NotSupportedException("Access native persistence is not qualified; SaveOnDispose is unavailable.");
    }
}

/// <summary>Explicit save target and shared loss/conflict policy.</summary>
public sealed class AccessSaveOptions {
    /// <summary>Null retains the document family, or uses a recognized destination extension.</summary>
    public AccessFileFormat? Format { get; set; }
    /// <summary>Optional explicit target generation. It must agree with the destination family.</summary>
    public AccessFormatProfile? Profile { get; set; }
    /// <summary>Conversion loss policy. Unsupported codecs remain blocked under every policy.</summary>
    public OfficeConversionLossPolicy LossPolicy { get; set; } = OfficeConversionLossPolicy.Block;
    /// <summary>Existing-file policy for a qualified writer.</summary>
    public OfficeConversionFileConflictPolicy FileConflictPolicy { get; set; } = OfficeConversionFileConflictPolicy.FailIfExists;
    internal void Validate() {
        if (Format.HasValue && !Enum.IsDefined(typeof(AccessFileFormat), Format.Value)) throw new ArgumentOutOfRangeException(nameof(Format));
        if (Profile.HasValue && (!Enum.IsDefined(typeof(AccessFormatProfile), Profile.Value) || Profile.Value == AccessFormatProfile.Unknown)) throw new ArgumentOutOfRangeException(nameof(Profile));
        if (!Enum.IsDefined(typeof(OfficeConversionLossPolicy), LossPolicy)) throw new ArgumentOutOfRangeException(nameof(LossPolicy));
        if (!Enum.IsDefined(typeof(OfficeConversionFileConflictPolicy), FileConflictPolicy)) throw new ArgumentOutOfRangeException(nameof(FileConflictPolicy));
    }
}
