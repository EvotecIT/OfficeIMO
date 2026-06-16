namespace OfficeIMO.Rtf;

/// <summary>
/// Document-level or section-level footnote and endnote numbering settings.
/// </summary>
public sealed class RtfNoteSettings {
    /// <summary>Beginning footnote number.</summary>
    public int? FootnoteStartNumber { get; set; }

    /// <summary>Footnote numbering restart behavior.</summary>
    public RtfNoteNumberRestart? FootnoteRestart { get; set; }

    /// <summary>Footnote number display format.</summary>
    public RtfNoteNumberFormat? FootnoteNumberFormat { get; set; }

    /// <summary>Footnote placement in the document or section.</summary>
    public RtfFootnotePlacement? FootnotePlacement { get; set; }

    /// <summary>Beginning endnote number.</summary>
    public int? EndnoteStartNumber { get; set; }

    /// <summary>Endnote numbering restart behavior.</summary>
    public RtfNoteNumberRestart? EndnoteRestart { get; set; }

    /// <summary>Endnote number display format.</summary>
    public RtfNoteNumberFormat? EndnoteNumberFormat { get; set; }

    /// <summary>Endnote placement in the document or section.</summary>
    public RtfEndnotePlacement? EndnotePlacement { get; set; }

    /// <summary>Sets footnote numbering controls.</summary>
    public RtfNoteSettings SetFootnoteNumbering(
        int? start = null,
        RtfNoteNumberRestart? restart = null,
        RtfNoteNumberFormat? format = null) {
        ValidatePositive(start, nameof(start));
        FootnoteStartNumber = start;
        FootnoteRestart = restart;
        FootnoteNumberFormat = format;
        return this;
    }

    /// <summary>Sets footnote placement.</summary>
    public RtfNoteSettings SetFootnotePlacement(RtfFootnotePlacement? placement) {
        FootnotePlacement = placement;
        return this;
    }

    /// <summary>Sets endnote numbering controls.</summary>
    public RtfNoteSettings SetEndnoteNumbering(
        int? start = null,
        RtfNoteNumberRestart? restart = null,
        RtfNoteNumberFormat? format = null) {
        ValidatePositive(start, nameof(start));
        EndnoteStartNumber = start;
        EndnoteRestart = restart;
        EndnoteNumberFormat = format;
        return this;
    }

    /// <summary>Sets endnote placement.</summary>
    public RtfNoteSettings SetEndnotePlacement(RtfEndnotePlacement? placement) {
        EndnotePlacement = placement;
        return this;
    }

    internal bool HasAnyValue =>
        FootnoteStartNumber.HasValue ||
        FootnoteRestart.HasValue ||
        FootnoteNumberFormat.HasValue ||
        FootnotePlacement.HasValue ||
        EndnoteStartNumber.HasValue ||
        EndnoteRestart.HasValue ||
        EndnoteNumberFormat.HasValue ||
        EndnotePlacement.HasValue;

    internal void Clear() {
        FootnoteStartNumber = null;
        FootnoteRestart = null;
        FootnoteNumberFormat = null;
        FootnotePlacement = null;
        EndnoteStartNumber = null;
        EndnoteRestart = null;
        EndnoteNumberFormat = null;
        EndnotePlacement = null;
    }

    private static void ValidatePositive(int? value, string parameterName) {
        if (value.HasValue && value.Value <= 0) {
            throw new ArgumentOutOfRangeException(parameterName, "Note number start must be greater than zero.");
        }
    }
}
