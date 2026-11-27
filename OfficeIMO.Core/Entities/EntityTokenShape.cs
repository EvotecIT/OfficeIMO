using System;

namespace OfficeIMO.Entities;

/// <summary>Token shapes a recognizer can request.</summary>
[Flags]
public enum EntityTokenShape {
    /// <summary>No token shapes.</summary>
    None = 0,
    /// <summary>A maximal ASCII digit run, including single separators.</summary>
    DigitRun = 1,
    /// <summary>An ASCII alphanumeric atom at a word boundary.</summary>
    AlphaNumericRun = 2,
    /// <summary>An at-sign anchor for email addresses.</summary>
    At = 4,
    /// <summary>A potentially supported URL scheme or www prefix.</summary>
    Scheme = 8,
}
