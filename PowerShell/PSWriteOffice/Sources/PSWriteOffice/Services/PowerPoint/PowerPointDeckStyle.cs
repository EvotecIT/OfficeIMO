namespace PSWriteOffice.Services.PowerPoint;

/// <summary>Optional composition tuning layered over the chosen document theme.</summary>
public enum PowerPointDeckStyle {
    /// <summary>Use the theme and any purpose-matched recipe.</summary>
    Auto,
    /// <summary>Emphasize visual evidence and proof panels.</summary>
    FieldProof,
    /// <summary>Use relaxed spacing for narrative case studies.</summary>
    EditorialCaseStudy,
    /// <summary>Use compact geometry for architecture and operational detail.</summary>
    TechnicalMap,
    /// <summary>Use dense supporting content with restrained presentation.</summary>
    QuietAppendix
}
