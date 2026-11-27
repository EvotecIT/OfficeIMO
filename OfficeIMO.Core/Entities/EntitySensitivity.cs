namespace OfficeIMO.Entities;

/// <summary>Entity classifications used by masking policies.</summary>
public enum EntitySensitivity {
    /// <summary>Personal data.</summary>
    Personal,
    /// <summary>Financial data.</summary>
    Financial,
    /// <summary>Organization-specific data.</summary>
    Organizational,
    /// <summary>Public-use identifiers and addresses.</summary>
    Public,
}
