namespace OfficeIMO.Toml;

/// <summary>
/// Raised when a JSON Patch or JSON Merge Patch cannot be applied: the patch is not valid JSON or is malformed, a
/// <c>test</c> fails, a pointer addresses nothing, a value is null (TOML has none), or the document refuses an edit. The
/// document is unchanged when it is raised.
/// </summary>
public sealed class TomlPatchException : Exception {
    internal TomlPatchException(int operationIndex, string? path, string message, Exception? innerException)
        : base(operationIndex < 0 ? message : "Patch operation " + operationIndex.ToString(CultureInfo.InvariantCulture) + (path is null ? string.Empty : " at '" + path + "'") + " failed: " + message, innerException) {
        OperationIndex = operationIndex;
        Path = path;
    }

    /// <summary>Gets the zero-based index of the JSON Patch operation that failed, or -1 when the patch as a whole did.</summary>
    public int OperationIndex { get; }

    /// <summary>Gets the <c>path</c> member of the operation that failed, when it has one.</summary>
    public string? Path { get; }
}
