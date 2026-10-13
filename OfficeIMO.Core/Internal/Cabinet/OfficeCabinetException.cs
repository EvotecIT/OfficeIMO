using System;
using System.IO;

namespace OfficeIMO.Core.Internal.Cabinet;

/// <summary>Describes a malformed or unsupported Cabinet record without imposing a document-specific diagnostic prefix.</summary>
internal sealed class OfficeCabinetException : IOException {
    internal OfficeCabinetException(string code, string message, long? offset = null, Exception? innerException = null)
        : base(message, innerException) {
        Code = code;
        Offset = offset;
    }

    internal string Code { get; }
    internal long? Offset { get; }
}
