namespace OfficeIMO.Rtf;

public sealed partial class RtfParagraph {
    // Conversion views isolate mutable formatting while retaining original content and resource identities.
    /// <summary>Copies paragraph controls without sharing visible content or a cached list marker.</summary>
    internal RtfParagraph CopyEmptyFormattingView() {
        RtfParagraph copy = CopyFormattingView();
        copy._runs = new List<RtfRun>();
        copy._inlines = new List<IRtfInline>();
        copy.ListText = null;
        return copy;
    }

    internal RtfParagraph CopyFormattingView() {
        var copy = (RtfParagraph)MemberwiseClone();
        var context = new RtfCloneContext();
        copy._tabStops = context.CloneList(_tabStops);
        copy.TopBorder = context.Clone(TopBorder)!;
        copy.LeftBorder = context.Clone(LeftBorder)!;
        copy.BottomBorder = context.Clone(BottomBorder)!;
        copy.RightBorder = context.Clone(RightBorder)!;
        copy.Frame = context.Clone(Frame)!;
        copy.LegacyNumbering = context.Clone(LegacyNumbering)!;
        return copy;
    }
}

public sealed partial class RtfRun {
    internal RtfRun CopyFormattingView() {
        var copy = (RtfRun)MemberwiseClone();
        copy.CharacterBorder = new RtfCloneContext().Clone(CharacterBorder)!;
        return copy;
    }
}
