using System;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Closes a PowerPoint presentation and optionally saves it.</summary>
/// <para>Provides a cmdlet wrapper so PowerShell scripts do not need to call <c>Dispose</c> directly.</para>
/// <example>
///   <summary>Close without saving.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$ppt = Get-OfficePowerPoint -Path .\deck.pptx; Close-OfficePowerPoint -Presentation $ppt</code>
///   <para>Releases the loaded presentation instance.</para>
/// </example>
/// <example>
///   <summary>Save, open, and close.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Close-OfficePowerPoint -Presentation $ppt -Save -Open</code>
///   <para>Saves the presentation, opens it in PowerPoint, and releases the object.</para>
/// </example>
[Cmdlet(VerbsCommon.Close, "OfficePowerPoint", SupportsShouldProcess = true)]
public sealed class CloseOfficePowerPointCommand : PSCmdlet {
    /// <summary>Presentation to close.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    [ValidateNotNull]
    public PowerPointPresentation Presentation { get; set; } = null!;

    /// <summary>Persist changes before closing.</summary>
    [Parameter]
    public SwitchParameter Save { get; set; }

    /// <summary>Optional target path when saving.</summary>
    [Parameter]
    [Alias("FilePath")]
    public string? Path { get; set; }

    /// <summary>Open the presentation after saving. Requires -Save or -Path.</summary>
    [Parameter]
    [Alias("Show")]
    public SwitchParameter Open { get; set; }

    /// <summary>Password used to save the presentation as an encrypted package.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (Presentation == null) {
            WriteError(new ErrorRecord(new ArgumentNullException(nameof(Presentation)), "PresentationNull", ErrorCategory.InvalidArgument, null));
            return;
        }

        if (Open.IsPresent && !Save.IsPresent && string.IsNullOrWhiteSpace(Path)) {
            throw new PSArgumentException("Use -Save or -Path with -Open so the presentation is persisted before it is opened.", nameof(Open));
        }

        try {
            var shouldSave = Save.IsPresent || !string.IsNullOrWhiteSpace(Path);
            var action = shouldSave ? "Save and close" : "Close";
            if (ShouldProcess("PowerPoint presentation", action)) {
                var resolvedPath = !string.IsNullOrWhiteSpace(Path)
                    ? SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path)
                    : null;
                PowerPointDocumentService.ClosePresentation(Presentation, shouldSave, Open.IsPresent, Password, resolvedPath);
            }
        } catch (Exception ex) {
            WriteError(new ErrorRecord(ex, "PowerPointCloseFailed", ErrorCategory.InvalidOperation, Presentation));
        }
    }
}
