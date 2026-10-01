namespace OfficeIMO.Rtf;

public sealed partial class RtfCharacterBorder : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfCharacterBorder)MemberwiseClone();
        context.Register(this, clone);
        return clone;
    }
}

public sealed partial class RtfLegacyNumbering : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfLegacyNumbering)MemberwiseClone();
        context.Register(this, clone);
        return clone;
    }
}

public sealed partial class RtfLineNumbering : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfLineNumbering)MemberwiseClone();
        context.Register(this, clone);
        return clone;
    }
}

public sealed partial class RtfNoteSettings : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfNoteSettings)MemberwiseClone();
        context.Register(this, clone);
        return clone;
    }
}

public sealed partial class RtfPageBorder : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfPageBorder)MemberwiseClone();
        context.Register(this, clone);
        return clone;
    }
}

public sealed partial class RtfPageBorders : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfPageBorders)MemberwiseClone();
        context.Register(this, clone);
        clone.Top = context.Clone(Top)!;
        clone.Bottom = context.Clone(Bottom)!;
        clone.Left = context.Clone(Left)!;
        clone.Right = context.Clone(Right)!;
        return clone;
    }
}

public sealed partial class RtfPageSetup : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfPageSetup)MemberwiseClone();
        context.Register(this, clone);
        clone.PageBorders = context.Clone(PageBorders)!;
        return clone;
    }
}

public sealed partial class RtfParagraphBorder : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfParagraphBorder)MemberwiseClone();
        context.Register(this, clone);
        return clone;
    }
}

public sealed partial class RtfStyle : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfStyle)MemberwiseClone();
        context.Register(this, clone);
        clone._tabStops = context.CloneList(_tabStops);
        clone.KeyCode = context.Clone(KeyCode)!;
        clone.TopBorder = context.Clone(TopBorder)!;
        clone.LeftBorder = context.Clone(LeftBorder)!;
        clone.BottomBorder = context.Clone(BottomBorder)!;
        clone.RightBorder = context.Clone(RightBorder)!;
        clone.LegacyNumbering = context.Clone(LegacyNumbering)!;
        clone.Frame = context.Clone(Frame)!;
        clone.TableRowFormat = context.Clone(TableRowFormat)!;
        return clone;
    }
}

public sealed partial class RtfStyleKeyCode : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfStyleKeyCode)MemberwiseClone();
        context.Register(this, clone);
        return clone;
    }
}

public sealed partial class RtfTableCellBorder : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfTableCellBorder)MemberwiseClone();
        context.Register(this, clone);
        return clone;
    }
}

public sealed partial class RtfTableRowBorder : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfTableRowBorder)MemberwiseClone();
        context.Register(this, clone);
        return clone;
    }
}

public sealed partial class RtfTabStop : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfTabStop)MemberwiseClone();
        context.Register(this, clone);
        return clone;
    }
}
