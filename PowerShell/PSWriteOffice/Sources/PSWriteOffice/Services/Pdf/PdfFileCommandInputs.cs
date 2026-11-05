using System;
using System.Collections.Generic;
using System.IO;
using System.Management.Automation;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Services.Pdf;

/// <summary>Resolves PowerShell PDF batches before native operations and asynchronous work.</summary>
internal static class PdfFileCommandInputs {
    internal static IReadOnlyList<(string Input, string Output)> Resolve(PSCmdlet command,
        string[]? path, string[]? literalPath, string? outputPath, string? outputDirectory, bool noClobber, string? outputExtension = null) {
        var paths = OfficePath.ResolveInputs(command, literalPath ?? path ?? Array.Empty<string>(), literalPath != null);
        if (outputPath != null && paths.Count != 1) throw new PSArgumentException("OutputPath requires exactly one input. Use OutputDirectory for a batch.");
        var result = new List<(string Input, string Output)>();
        foreach (var input in paths) {
            string output = outputPath != null ? OfficePath.ResolveOutput(command.SessionState, outputPath, false)
                : Path.Combine(OfficePath.ResolveOutput(command.SessionState, outputDirectory!, true), outputExtension == null ? Path.GetFileName(input) : Path.ChangeExtension(Path.GetFileName(input), outputExtension));
            OfficePath.ValidateDestination(output, noClobber);
            result.Add((input, output));
        }
        if (result.SelectOutputCollision()) throw new PSArgumentException("Two input files have the same output name. Convert them separately or use distinct output folders.");
        return result;
    }
    private static bool SelectOutputCollision(this List<(string Input, string Output)> paths) {
        var outputs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var pair in paths) if (!outputs.Add(pair.Output)) return true;
        return false;
    }
}
