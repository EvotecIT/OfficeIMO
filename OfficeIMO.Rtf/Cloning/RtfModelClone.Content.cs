namespace OfficeIMO.Rtf;

public sealed partial class RtfBookmarkMarker : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfBookmarkMarker)MemberwiseClone();
        context.Register(this, clone);
        return clone;
    }
}

public sealed partial class RtfBreak : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfBreak)MemberwiseClone();
        context.Register(this, clone);
        return clone;
    }
}

public sealed partial class RtfDocument : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfDocument)MemberwiseClone();
        context.Register(this, clone);
        clone._paragraphs = context.CloneList(_paragraphs);
        clone._blocks = context.CloneList(_blocks);
        clone._fonts = context.CloneList(_fonts);
        clone._colors = context.CloneList(_colors);
        clone._styles = context.CloneList(_styles);
        clone._listDefinitions = context.CloneList(_listDefinitions);
        clone._listOverrides = context.CloneList(_listOverrides);
        clone._headerFooters = context.CloneList(_headerFooters);
        clone._notes = context.CloneList(_notes);
        clone._sections = context.CloneList(_sections);
        clone._userProperties = context.CloneList(_userProperties);
        clone._documentVariables = context.CloneList(_documentVariables);
        clone._revisionAuthors = context.CloneList(_revisionAuthors);
        clone._revisionSaveIds = context.CloneList(_revisionSaveIds);
        clone._fileReferences = context.CloneList(_fileReferences);
        clone._xmlNamespaces = context.CloneList(_xmlNamespaces);
        clone.Info = context.Clone(Info)!;
        clone.PageSetup = context.Clone(PageSetup)!;
        clone.Settings = context.Clone(Settings)!;
        clone.NoteSettings = context.Clone(NoteSettings)!;
        return clone;
    }
}

public sealed partial class RtfField : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfField)MemberwiseClone();
        context.Register(this, clone);
        clone.Result = context.Clone(Result)!;
        clone.HyperlinkField = context.Clone(HyperlinkField)!;
        clone.FormFieldData = context.Clone(FormFieldData)!;
        return clone;
    }
}

public sealed partial class RtfGeneratedText : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfGeneratedText)MemberwiseClone();
        context.Register(this, clone);
        clone.Note = context.Clone(Note)!;
        return clone;
    }
}

public sealed partial class RtfHeaderFooter : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfHeaderFooter)MemberwiseClone();
        context.Register(this, clone);
        clone._paragraphs = context.CloneList(_paragraphs);
        return clone;
    }
}

public sealed partial class RtfHyperlinkFieldInfo : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfHyperlinkFieldInfo)MemberwiseClone();
        context.Register(this, clone);
        return clone;
    }
}

public sealed partial class RtfImage : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfImage)MemberwiseClone();
        context.Register(this, clone);
        clone.Data = context.Clone(Data)!;
        return clone;
    }
}

public sealed partial class RtfNote : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfNote)MemberwiseClone();
        context.Register(this, clone);
        clone._paragraphs = context.CloneList(_paragraphs);
        return clone;
    }
}

public sealed partial class RtfObject : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfObject)MemberwiseClone();
        context.Register(this, clone);
        clone.Data = context.Clone(Data)!;
        clone.Result = context.Clone(Result)!;
        clone.ResultImage = context.Clone(ResultImage)!;
        return clone;
    }
}

public sealed partial class RtfParagraph : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfParagraph)MemberwiseClone();
        context.Register(this, clone);
        clone._runs = context.CloneList(_runs);
        clone._inlines = context.CloneList(_inlines);
        clone._tabStops = context.CloneList(_tabStops);
        clone.LegacyNumbering = context.Clone(LegacyNumbering)!;
        clone.ListText = context.Clone(ListText)!;
        clone.TopBorder = context.Clone(TopBorder)!;
        clone.LeftBorder = context.Clone(LeftBorder)!;
        clone.BottomBorder = context.Clone(BottomBorder)!;
        clone.RightBorder = context.Clone(RightBorder)!;
        clone.Frame = context.Clone(Frame)!;
        return clone;
    }
}

public sealed partial class RtfParagraphFrame : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfParagraphFrame)MemberwiseClone();
        context.Register(this, clone);
        return clone;
    }
}

public sealed partial class RtfRun : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfRun)MemberwiseClone();
        context.Register(this, clone);
        clone.CharacterBorder = context.Clone(CharacterBorder)!;
        clone.Note = context.Clone(Note)!;
        return clone;
    }
}

public sealed partial class RtfSection : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfSection)MemberwiseClone();
        context.Register(this, clone);
        clone._blocks = context.CloneList(_blocks);
        clone._columns = context.CloneList(_columns);
        clone._document = context.Clone(_document)!;
        clone.PageSetup = context.Clone(PageSetup)!;
        clone.NoteSettings = context.Clone(NoteSettings)!;
        clone.LineNumbering = context.Clone(LineNumbering)!;
        return clone;
    }
}

public sealed partial class RtfSectionColumn : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfSectionColumn)MemberwiseClone();
        context.Register(this, clone);
        return clone;
    }
}

public sealed partial class RtfShape : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfShape)MemberwiseClone();
        context.Register(this, clone);
        clone._instructions = context.CloneList(_instructions);
        clone._properties = context.CloneList(_properties);
        clone._textBoxParagraphs = context.CloneList(_textBoxParagraphs);
        return clone;
    }
}

public sealed partial class RtfShapeInstruction : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfShapeInstruction)MemberwiseClone();
        context.Register(this, clone);
        return clone;
    }
}

public sealed partial class RtfShapeProperty : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfShapeProperty)MemberwiseClone();
        context.Register(this, clone);
        return clone;
    }
}

public sealed partial class RtfTable : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfTable)MemberwiseClone();
        context.Register(this, clone);
        clone._rows = context.CloneList(_rows);
        return clone;
    }
}

public sealed partial class RtfTableCell : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfTableCell)MemberwiseClone();
        context.Register(this, clone);
        clone._paragraphs = context.CloneList(_paragraphs);
        clone._blocks = context.CloneList(_blocks);
        clone.TopBorder = context.Clone(TopBorder)!;
        clone.LeftBorder = context.Clone(LeftBorder)!;
        clone.BottomBorder = context.Clone(BottomBorder)!;
        clone.RightBorder = context.Clone(RightBorder)!;
        clone.TopLeftToBottomRightBorder = context.Clone(TopLeftToBottomRightBorder)!;
        clone.TopRightToBottomLeftBorder = context.Clone(TopRightToBottomLeftBorder)!;
        return clone;
    }
}

public sealed partial class RtfTableRow : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfTableRow)MemberwiseClone();
        context.Register(this, clone);
        clone._cells = context.CloneList(_cells);
        clone.TopBorder = context.Clone(TopBorder)!;
        clone.LeftBorder = context.Clone(LeftBorder)!;
        clone.BottomBorder = context.Clone(BottomBorder)!;
        clone.RightBorder = context.Clone(RightBorder)!;
        clone.HorizontalBorder = context.Clone(HorizontalBorder)!;
        clone.VerticalBorder = context.Clone(VerticalBorder)!;
        return clone;
    }
}
