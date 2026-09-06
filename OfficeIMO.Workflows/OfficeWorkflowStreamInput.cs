namespace OfficeIMO.Workflows;

/// <summary>A reopenable provider input whose original location remains in the workflow request.</summary>
public sealed class OfficeWorkflowStreamInput {
    /// <summary>Creates an input. Each invocation must return a new readable stream; the runner closes it.</summary>
    /// <param name="name">The provider's display filename, including its format extension.</param>
    /// <param name="openRead">Opens the selected document with the provider's permission scope.</param>
    /// <param name="expectedSha256">Optional hexadecimal SHA-256 captured when the input was selected.</param>
    public OfficeWorkflowStreamInput(string name, Func<CancellationToken, Task<Stream>> openRead, string? expectedSha256 = null) {
        if (string.IsNullOrWhiteSpace(name) || name.Length > 4096 || name.IndexOfAny(['/', '\\', '\0']) >= 0) {
            throw new ArgumentException("A provider filename is required.", nameof(name));
        }
        if (expectedSha256 is not null && (expectedSha256.Length != 64 || expectedSha256.Any(character => !Uri.IsHexDigit(character)))) {
            throw new ArgumentException("The expected fingerprint must be a hexadecimal SHA-256.", nameof(expectedSha256));
        }
        Name = name;
        OpenRead = openRead ?? throw new ArgumentNullException(nameof(openRead));
        ExpectedSha256 = expectedSha256;
    }

    /// <summary>Gets the selected filename. This name determines format routing, never a staging directory.</summary>
    public string Name { get; }
    /// <summary>Gets the stream factory; the runner invokes it again to verify source contents before publication.</summary>
    public Func<CancellationToken, Task<Stream>> OpenRead { get; }
    /// <summary>Gets the optional fingerprint captured when the input was selected.</summary>
    public string? ExpectedSha256 { get; }
}
