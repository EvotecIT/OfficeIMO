using System.IO;

namespace OfficeIMO;

/// <summary>Internal compression failure translated by the format-specific adapter.</summary>
internal sealed class OfficeLzxException : IOException {
    internal OfficeLzxException(string code, string message) : base(message) { Code = code; }
    internal string Code { get; }
}
