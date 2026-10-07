using OfficeIMO.Pdf;
using OfficeIMO.Workflows;

namespace OfficeIMO.Tool.Agent;

/// <summary>Bounded surface policy shared by CLI and MCP adapters over the reusable workflows.</summary>
internal sealed class PdfWorkflowSettings {
    internal string? Pages { get; init; }
    internal string? PasswordEnvironmentVariable { get; init; }
    internal bool Overwrite { get; init; }
    internal bool AcknowledgeRasterOutput { get; init; }
    internal int MaximumPages { get; init; } = 100;
    internal double Dpi { get; init; } = 150;
    internal long MaximumPixelsPerPage { get; init; } = 25_000_000;
    internal long MaximumInputBytes { get; init; } = 256L * 1024 * 1024;
    internal long MaximumOutputBytes { get; init; } = 512L * 1024 * 1024;

    internal void Validate() {
        if (MaximumPages is < 1 or > 10_000) throw new AgentUsageException("maximumPages must be from 1 through 10000.");
        if (!double.IsFinite(Dpi) || Dpi is < 36 or > 600) throw new AgentUsageException("dpi must be from 36 through 600.");
        if (MaximumPixelsPerPage is < 1 or > 100_000_000) throw new AgentUsageException("maximumPixelsPerPage must be from 1 through 100000000.");
        if (MaximumInputBytes is < 1 or > 256L * 1024 * 1024) throw new AgentUsageException("maximumInputBytes must be from 1 through 268435456.");
        if (MaximumOutputBytes is < 1 or > 512L * 1024 * 1024) throw new AgentUsageException("maximumOutputBytes must be from 1 through 536870912.");
        _ = Selector();
    }

    internal PdfPageSelector? Selector() {
        if (Pages is null) return null;
        if (string.IsNullOrWhiteSpace(Pages) || Pages.Length > 4096) throw new AgentUsageException("pages must contain 1 through 4096 characters.");
        try { return PdfPageSelector.Parse(Pages); }
        catch (Exception error) when (error is ArgumentException or FormatException or OverflowException) {
            throw new AgentUsageException("Invalid PDF page selection.", error);
        }
    }

    internal OfficeWorkflowLimits Limits() => new() { MaximumInputBytes = MaximumInputBytes, MaximumOutputBytes = MaximumOutputBytes };
    internal OfficeWorkflowConflictPolicy ConflictPolicy => Overwrite ? OfficeWorkflowConflictPolicy.Replace : OfficeWorkflowConflictPolicy.Fail;

    internal string? Password() {
        if (PasswordEnvironmentVariable is null) return null;
        if (PasswordEnvironmentVariable.Length is < 1 or > 128 || !PasswordEnvironmentVariable.All(character => char.IsAsciiLetterOrDigit(character) || character == '_'))
            throw new AgentUsageException("passwordEnvironmentVariable must be a simple environment-variable name.");
        return Environment.GetEnvironmentVariable(PasswordEnvironmentVariable) is { Length: > 0 } secret
            ? secret : throw new AgentUsageException("The selected password environment variable is missing or empty.");
    }
}
