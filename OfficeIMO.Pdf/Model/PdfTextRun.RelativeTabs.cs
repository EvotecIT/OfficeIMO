namespace OfficeIMO.Pdf;

public sealed partial class PdfTextRun {
    /// <summary>A source positional tab follows its paragraph frame when the available width changes.</summary>
    internal (double Fraction, double Offset)? RelativeTabStop { get; private set; }

    internal PdfTextRun WithRelativeTabFrom(PdfTextRun source) {
        RelativeTabStop = source.RelativeTabStop;
        return this;
    }
    internal static PdfTextRun RelativeTab(PdfTabAlignment alignment, PdfTabLeaderStyle leader, double fraction, double offset) {
        if (fraction < 0D || fraction > 1D || double.IsNaN(fraction) || double.IsInfinity(fraction))
            throw new System.ArgumentOutOfRangeException(nameof(fraction));
        if (double.IsNaN(offset) || double.IsInfinity(offset))
            throw new System.ArgumentOutOfRangeException(nameof(offset));
        return new PdfTextRun("\t", tabLeader: leader, tabAlignment: alignment) { RelativeTabStop = (fraction, offset) };
    }
}
