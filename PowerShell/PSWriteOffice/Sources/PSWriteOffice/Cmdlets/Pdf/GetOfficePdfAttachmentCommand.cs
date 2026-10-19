using System.IO;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Gets or extracts embedded file attachments from a PDF.</summary>
/// <example>
///   <summary>List and extract embedded PDF attachments.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$proof = @(
///     Get-OfficePdfAttachment -Path .\Examples\Documents\PdfWithAttachment.pdf
///     Get-OfficePdfAttachment -Path .\Examples\Documents\PdfWithAttachment.pdf -OutputDirectory .\Examples\Documents\Attachments
/// )
/// $proof</code>
///   <para>First returns attachment metadata, then writes embedded files to disk.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficePdfAttachment")]
[OutputType(typeof(PdfExtractedAttachment), typeof(FileInfo))]
public sealed class GetOfficePdfAttachmentCommand : PSCmdlet
{
    /// <summary>PDF file path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Optional attachment name or file name filter.</summary>
    [Parameter]
    public string? Name { get; set; }

    /// <summary>Optional directory where attachments should be written.</summary>
    [Parameter]
    public string? OutputDirectory { get; set; }

    /// <summary>Password used to read a Standard password-encrypted PDF.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <summary>After successful password authentication, explicitly ignore owner-imposed extraction restrictions.</summary>
    [Parameter]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var attachments = PdfDocument.Load(
            PdfCommandUtilities.ResolvePath(this, Path),
            PdfCommandUtilities.CreateReadOptions(Password, IgnorePermissionRestrictions.IsPresent)).Attachments.Extract();
        string? outputDirectory = null;
        if (!string.IsNullOrWhiteSpace(OutputDirectory))
        {
            outputDirectory = PdfCommandUtilities.ResolvePath(this, OutputDirectory!);
            PdfCommandUtilities.EnsureOutputDirectory(outputDirectory);
        }

        foreach (var attachment in attachments)
        {
            if (!Matches(attachment))
            {
                continue;
            }

            if (outputDirectory == null)
            {
                WriteObject(attachment);
                continue;
            }

            var fileName = !string.IsNullOrWhiteSpace(attachment.UnicodeFileName)
                ? attachment.UnicodeFileName!
                : attachment.FileName;
            var outputPath = AtomicFileWriter.WriteUnique(outputDirectory, fileName, attachment.Bytes);
            WriteObject(new FileInfo(outputPath));
        }
    }

    private bool Matches(PdfExtractedAttachment attachment)
    {
        return string.IsNullOrWhiteSpace(Name) ||
            string.Equals(attachment.Name, Name, System.StringComparison.OrdinalIgnoreCase) ||
            string.Equals(attachment.FileName, Name, System.StringComparison.OrdinalIgnoreCase) ||
            string.Equals(attachment.UnicodeFileName, Name, System.StringComparison.OrdinalIgnoreCase);
    }
}
