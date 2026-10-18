using System.IO;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Draws arbitrary visual canvas content on existing PDF pages.</summary>
/// <remarks>
/// The script receives a <see cref="PdfPageCanvas"/> and <see cref="PdfStampPageContext"/>.
/// It can draw text, rich text, images, shapes, drawings, and tables. Interactive annotations,
/// links, form fields, and outlines are separate PDF operations and are rejected by this visual-only surface.
/// </remarks>
/// <example>
///   <summary>Draw a page-aware review band.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Add-OfficePdfCanvas -Path .\Report.pdf -OutputPath .\Stamped.pdf -PageRange '1,last' -Content {
///     param($canvas, $page)
///     $null = $canvas.Text("Review copy $($page.PageNumber)/$($page.PageCount)", 36, 24, $page.Width - 72, 24, 10)
/// }</code>
///   <para>The callback runs once for every selected page and may mix any supported visual canvas primitives.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficePdfCanvas", SupportsShouldProcess = true)]
[Alias("PdfCanvasStamp")]
[OutputType(typeof(FileInfo))]
public sealed class AddOfficePdfCanvasCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    /// <summary>Input PDF path.</summary>
    [Parameter(Mandatory = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Output PDF path.</summary>
    [Parameter(Mandatory = true)]
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>Canvas callback. Declare parameters for the canvas and page context.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public ScriptBlock Content { get; set; } = null!;

    /// <summary>Target page selector such as 1-3,odd,last. Omit to stamp every page.</summary>
    [Parameter]
    public string? PageRange { get; set; }

    /// <summary>Place the generated canvas behind existing page content.</summary>
    [Parameter]
    public SwitchParameter BehindContent { get; set; }

    /// <summary>Opacity applied to the complete generated canvas.</summary>
    [Parameter]
    [ValidateRange(0D, 1D)]
    public double Opacity { get; set; } = 1D;

    /// <summary>Configures native generated-PDF rendering options for canvas content, including embedded fonts and text shaping.</summary>
    /// <remarks>The callback receives a <see cref="PdfOptions"/> instance. Page geometry, margins, and encryption remain controlled by the stamping operation.</remarks>
    [Parameter]
    public ScriptBlock? ConfigureRendering { get; set; }

    /// <summary>Password used to authenticate an encrypted input PDF.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <summary>
    /// After successful password authentication, explicitly ignore owner-imposed usage restrictions.
    /// This does not discover, bypass, or crack a missing password.
    /// </summary>
    [Parameter]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var outputPath = PdfCommandUtilities.ResolvePath(this, OutputPath);
        if (!PdfCommandUtilities.ShouldWrite(this, outputPath, "Write canvas-stamped PDF")) {
            return;
        }

        var readOptions = PdfCommandUtilities.CreateReadOptions(Password, IgnorePermissionRestrictions.IsPresent);
        var document = PdfDocument.Load(PdfCommandUtilities.ResolvePath(this, Path), readOptions);
        PdfOptions? renderingOptions = null;
        if (ConfigureRendering is not null) {
            renderingOptions = new PdfOptions();
            _ = ConfigureRendering.Invoke(renderingOptions);
        }

        var options = new PdfCanvasStampOptions {
            BehindContent = BehindContent.IsPresent,
            Opacity = Opacity,
            RenderingOptions = renderingOptions
        };
        if (!string.IsNullOrWhiteSpace(PageRange)) {
            options.UseTargetPages(PageRange!);
        }

        var result = document.Stamp.Content(
            (canvas, page) => {
                using (PdfCanvasDslContext.Enter(canvas, page)) {
                    _ = Content.Invoke(canvas, page);
                }
            },
            options,
            readOptions);
        PdfCommandUtilities.EnsureDirectory(outputPath);
        result.Save(outputPath).RequireSuccess();
        WritePassThru(new FileInfo(outputPath));
    }
}
