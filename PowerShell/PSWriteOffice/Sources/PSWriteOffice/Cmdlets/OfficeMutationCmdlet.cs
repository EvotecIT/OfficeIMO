using System.IO;
using System.Management.Automation;

namespace PSWriteOffice.Cmdlets;

/// <summary>Base class for mutating commands that are quiet unless <c>-PassThru</c> is requested.</summary>
public abstract class OfficeMutationCmdlet : PSCmdlet {
    /// <summary>Emit the object created or changed by the command.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <summary>Writes the mutated object only when <see cref="PassThru"/> is present.</summary>
    /// <param name="value">Object created or changed by the command.</param>
    protected void WritePassThru(object? value) {
        if (PassThru.IsPresent && value != null) {
            WriteObject(value);
        }
    }

    /// <summary>Emits a saved file for an owned path-based document, or the still-live document otherwise.</summary>
    /// <param name="document">Document changed by the command.</param>
    /// <param name="ownsDocument">Whether the command opened and will dispose the document.</param>
    /// <param name="inputPath">Path used to open an owned document.</param>
    protected void WritePassThru(object document, bool ownsDocument, string? inputPath) {
        if (!PassThru.IsPresent) {
            return;
        }

        if (!ownsDocument) {
            WriteObject(document);
            return;
        }

        var path = SessionState.Path.GetUnresolvedProviderPathFromPSPath(inputPath ?? string.Empty);
        WriteObject(new FileInfo(path));
    }
}