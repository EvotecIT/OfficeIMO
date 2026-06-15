namespace OfficeIMO.Rtf;

/// <summary>
/// Character emitted after an RTF list level marker.
/// </summary>
public enum RtfListLevelFollowCharacter {
    /// <summary>A tab follows the list marker.</summary>
    Tab,

    /// <summary>A space follows the list marker.</summary>
    Space,

    /// <summary>No automatic character follows the list marker.</summary>
    Nothing
}
