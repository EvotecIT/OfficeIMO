namespace OfficeIMO.Rtf;

public sealed partial class RtfColor : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfColor)MemberwiseClone();
        context.Register(this, clone);
        return clone;
    }
}

public sealed partial class RtfDocumentInfo : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfDocumentInfo)MemberwiseClone();
        context.Register(this, clone);
        return clone;
    }
}

public sealed partial class RtfDocumentSettings : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfDocumentSettings)MemberwiseClone();
        context.Register(this, clone);
        return clone;
    }
}

public sealed partial class RtfDocumentVariable : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfDocumentVariable)MemberwiseClone();
        context.Register(this, clone);
        return clone;
    }
}

public sealed partial class RtfFileReference : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfFileReference)MemberwiseClone();
        context.Register(this, clone);
        return clone;
    }
}

public sealed partial class RtfFont : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfFont)MemberwiseClone();
        context.Register(this, clone);
        clone.Embedding = context.Clone(Embedding)!;
        return clone;
    }
}

public sealed partial class RtfFontEmbedding : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfFontEmbedding)MemberwiseClone();
        context.Register(this, clone);
        clone.Data = context.Clone(Data)!;
        return clone;
    }
}

public sealed partial class RtfFormFieldData : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfFormFieldData)MemberwiseClone();
        context.Register(this, clone);
        clone._controls = context.CloneList(_controls);
        clone._dropDownItems = context.CloneList(_dropDownItems);
        return clone;
    }
}

public sealed partial class RtfFormFieldDataControl : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfFormFieldDataControl)MemberwiseClone();
        context.Register(this, clone);
        return clone;
    }
}

public sealed partial class RtfListDefinition : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfListDefinition)MemberwiseClone();
        context.Register(this, clone);
        clone._levels = context.CloneList(_levels);
        return clone;
    }
}

public sealed partial class RtfListLevel : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfListLevel)MemberwiseClone();
        context.Register(this, clone);
        return clone;
    }
}

public sealed partial class RtfListLevelOverride : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfListLevelOverride)MemberwiseClone();
        context.Register(this, clone);
        return clone;
    }
}

public sealed partial class RtfListOverride : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfListOverride)MemberwiseClone();
        context.Register(this, clone);
        clone._levelOverrides = context.CloneList(_levelOverrides);
        return clone;
    }
}

public sealed partial class RtfRevisionAuthor : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfRevisionAuthor)MemberwiseClone();
        context.Register(this, clone);
        return clone;
    }
}

public sealed partial class RtfUserProperty : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfUserProperty)MemberwiseClone();
        context.Register(this, clone);
        return clone;
    }
}

public sealed partial class RtfXmlNamespace : IRtfCloneable {
    object IRtfCloneable.CloneModel(RtfCloneContext context) {
        var clone = (RtfXmlNamespace)MemberwiseClone();
        context.Register(this, clone);
        return clone;
    }
}
