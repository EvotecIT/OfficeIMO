using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Pdf;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Creates a PowerPoint presentation using the DSL.</summary>
/// <para>Initializes a presentation, runs the DSL script block, and optionally saves the deck.</para>
/// <example>
///   <summary>Create and capture the presentation object.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$ppt = New-OfficePowerPoint -Path .\deck.pptx -NoSave</code>
///   <para>Creates a live presentation associated with <c>deck.pptx</c> for incremental composition.</para>
/// </example>
/// <example>
///   <summary>Create a deck with a title slide.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePowerPoint -Path .\deck.pptx { PptSlide { PptTitle -Title 'Status Update' } } -Open</code>
///   <para>Creates, saves, and opens a deck with one titled slide.</para>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficePowerPoint", SupportsShouldProcess = true)]
[Alias("OfficePowerPoint", "PowerPointNew", "PptNew")]
public class NewOfficePowerPointCommand : PSCmdlet {
    /// <summary>Destination path for the new .pptx.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>DSL scriptblock describing presentation content.</summary>
    [Parameter(Position = 1)]
    public ScriptBlock? Content { get; set; }

    /// <summary>Open the presentation after saving.</summary>
    [Parameter]
    public SwitchParameter Open { get; set; }

    /// <summary>Skip saving after executing the DSL.</summary>
    [Parameter]
    public SwitchParameter NoSave { get; set; }

    /// <summary>Emit the saved <see cref="FileInfo"/> for chaining.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <summary>Password used to save the presentation as an encrypted package.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (NoSave.IsPresent && Open.IsPresent) {
            throw new PSArgumentException("-Open cannot be used with -NoSave because no file is written. Save the returned presentation explicitly, then use -Open on Save-OfficePowerPoint.", nameof(Open));
        }

        var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
        if (NoSave.IsPresent) {
            if (!PdfCommandUtilities.ShouldWrite(this, resolvedPath, "Create PowerPoint presentation")) {
                return;
            }
        } else {
            if (!PdfCommandUtilities.ShouldWrite(this, resolvedPath, "Write new PowerPoint presentation")) {
                return;
            }
        }

        var directory = System.IO.Path.GetDirectoryName(resolvedPath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory)) {
            Directory.CreateDirectory(directory);
        }

        PowerPointPresentation? presentation = null;
        try {
            presentation = PowerPointDocumentService.CreatePresentation(resolvedPath);

            if (Content != null) {
                using (PowerPointDslContext.Enter(presentation)) {
                    try {
                        Content.InvokeReturnAsIs();
                    } catch (Exception ex) when (IsStopUpstream(ex)) {
                        // Select-Object -First throws this to stop enumeration; ignore.
                    }
                }
            }

            if (NoSave.IsPresent) {
                WriteObject(presentation);
                return;
            }

            var savedPath = PowerPointDocumentService.SavePresentation(
                presentation,
                show: false,
                Password,
                filePath: null);
            PowerPointDocumentService.ClosePresentation(presentation, save: false, show: false);
            presentation = null;

            if (Open.IsPresent) {
                FileOpenService.Open(savedPath);
            }

            if (PassThru.IsPresent) {
                WriteObject(new FileInfo(resolvedPath));
            }
        } catch (Exception ex) {
            if (presentation != null) {
                PowerPointDocumentService.ClosePresentation(presentation, save: false, show: false);
            }
            WriteError(new ErrorRecord(ex, "PowerPointCreateFailed", ErrorCategory.InvalidOperation, Path));
        }
    }

    private static bool IsStopUpstream(Exception ex) {
        return ex.GetType().Name == "StopUpstreamCommandsException";
    }

}
