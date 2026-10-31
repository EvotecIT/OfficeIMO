using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.OpenDocument;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.OpenDocument;

namespace PSWriteOffice.Cmdlets.OpenDocument;

/// <summary>Creates an ODT, ODS, ODP or ODG model, returning it live or saving a completed content block.</summary>
/// <para>Kind is inferred from Path when omitted. Without a block no file is written. A block publishes only after success; an existing destination requires Force.</para>
/// <example><summary>Create a text report and save the completed block.</summary>
/// <code>New-OfficeOpenDocument -Path ./Report.odt {
///     Add-OfficeOpenDocumentHeading -Text 'Service report' -Level 1
///     Add-OfficeOpenDocumentParagraph -Text 'All services are available.'
/// }</code></example>
/// <example><summary>Create a live drawing and save it explicitly.</summary>
/// <code>$drawing = New-OfficeOpenDocument -Path ./Overview.odg
/// $page = $drawing.AddPage('Overview')
/// Save-OfficeOpenDocument -Document $drawing</code></example>
[Cmdlet(VerbsCommon.New, "OfficeOpenDocument", SupportsShouldProcess = true)]
[OutputType(typeof(OdfDocument), typeof(FileInfo))]
public sealed class NewOfficeOpenDocumentCommand : OfficeLifecycleCmdlet<OdfDocument> {
    /// <summary>Optional associated destination; required with a content block.</summary>
    [Parameter(Position = 0), ValidateNotNullOrEmpty]
    public string? Path { get; set; }

    /// <summary>Document kind; required when no destination extension is supplied.</summary>
    [Parameter]
    public OdfDocumentKind? Kind { get; set; }

    /// <summary>Allow a completed content block to replace an existing destination.</summary>
    [Parameter]
    public SwitchParameter Force { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        OdfDocumentKind? kind = Kind;
        if (Path != null) {
            string destination = ResolveOutputPath(Path);
            if (!kind.HasValue) {
                kind = System.IO.Path.GetExtension(destination).ToLowerInvariant() switch {
                    ".odt" => OdfDocumentKind.Text,
                    ".ods" => OdfDocumentKind.Spreadsheet,
                    ".odp" => OdfDocumentKind.Presentation,
                    ".odg" => OdfDocumentKind.Graphics,
                    _ => throw new PSArgumentException("Path must use .odt, .ods, .odp or .odg, or supply the matching Kind.", nameof(Path))
                };
            }
            OpenDocumentCommandUtilities.ValidateOpenDocumentExtension(destination, kind.Value, nameof(Path));
        }
        if (!kind.HasValue) {
            InvalidArgument(new PSArgumentException("Supply Kind or a Path with an OpenDocument extension.", nameof(Kind)), Kind);
            return;
        }
        RunDocument(Path, isNew: true, force: Force.IsPresent,
            path => OdfDocument.Create(kind.Value, path),
            (document, path, allowLoss) => {
                OdfSaveResult result = document.Save(path);
                if (!allowLoss) result.Report.RequireNoLoss();
                return null;
            });
    }

    /// <inheritdoc />
    protected override void InvokeContent(OdfDocument document) {
        using (OpenDocumentDslContext.Enter(document)) Content!.InvokeReturnAsIs();
    }
}
