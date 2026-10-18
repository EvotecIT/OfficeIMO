using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Converts Word documents between supported .doc and .docx formats.</summary>
/// <para>Uses the OfficeIMO Word normal load/save conversion path, including legacy DOC diagnostics and save preflight.</para>
/// <example>
///   <summary>Convert a legacy DOC file to DOCX.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ConvertTo-OfficeWordDocument -Path .\legacy.doc -OutputPath .\converted.docx -PassThru</code>
///   <para>Reads the .doc file and writes a .docx file.</para>
/// </example>
/// <example>
///   <summary>Convert a DOCX file to native DOC.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ConvertTo-OfficeWordDocument -Path .\report.docx -OutputPath .\report.doc -Force</code>
///   <para>Writes a supported native Word 97-2003 .doc file.</para>
/// </example>
[Cmdlet(VerbsData.ConvertTo, "OfficeWordDocument", SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo))]
public sealed class ConvertToOfficeWordDocumentCommand : PSCmdlet
{
    /// <summary>Source .doc or .docx file path.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Destination .doc or .docx file path.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    [Alias("OutPath")]
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>Overwrite an existing destination file.</summary>
    [Parameter]
    public SwitchParameter Force { get; set; }

    /// <summary>Allow conversion when a legacy DOC source contains unsupported or preserve-only content.</summary>
    [Parameter]
    public SwitchParameter AllowLossyLegacyConversion { get; set; }

    /// <summary>Open the converted document after saving.</summary>
    [Parameter]
    public SwitchParameter Open { get; set; }

    /// <summary>Emit the saved file information.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        try
        {
            var sourcePath = ResolvePath(Path);
            var outputPath = ResolvePath(OutputPath);

            if (!File.Exists(sourcePath))
            {
                throw new FileNotFoundException($"File '{sourcePath}' was not found.", sourcePath);
            }

            if (File.Exists(outputPath) && !Force.IsPresent)
            {
                throw new IOException($"File '{outputPath}' already exists. Use -Force to overwrite it.");
            }

            var action = $"Convert Word document to {System.IO.Path.GetExtension(outputPath)}";
            if (!ShouldProcess(outputPath, action))
            {
                return;
            }

            PdfCommandUtilities.EnsureDirectory(outputPath);
            if (string.Equals(System.IO.Path.GetExtension(sourcePath), System.IO.Path.GetExtension(outputPath), StringComparison.OrdinalIgnoreCase))
            {
                File.Copy(sourcePath, outputPath, overwrite: Force.IsPresent);
            }
            else
            {
                WordDocument.Convert(sourcePath, outputPath, new WordDocumentConversionOptions
                {
                    FileConflictPolicy = Force.IsPresent
                        ? OfficeConversionFileConflictPolicy.Replace
                        : OfficeConversionFileConflictPolicy.FailIfExists,
                    LossPolicy = AllowLossyLegacyConversion.IsPresent
                        ? OfficeConversionLossPolicy.Allow
                        : OfficeConversionLossPolicy.Block
                });
            }

            if (Open.IsPresent)
            {
                FileOpenService.Open(outputPath);
            }

            if (PassThru.IsPresent)
            {
                WriteObject(new FileInfo(outputPath));
            }
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "ConvertToOfficeWordDocumentFailed", ErrorCategory.InvalidOperation, Path));
        }
    }

    private string ResolvePath(string path)
    {
        return PdfCommandUtilities.ResolvePath(this, path);
    }
}
