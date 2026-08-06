using System;

namespace OfficeIMO.Drawing;

/// <summary>Reports a stable failure while parsing portable mathematical markup.</summary>
public sealed class OfficeMathParseException : FormatException {
    /// <summary>Creates a parse exception with a stable code.</summary>
    public OfficeMathParseException(string code, string message) : base(message) {
        if (string.IsNullOrWhiteSpace(code)) throw new ArgumentException("A parse error code is required.", nameof(code));
        Code = code;
    }

    /// <summary>Stable machine-readable error code.</summary>
    public string Code { get; }
}
