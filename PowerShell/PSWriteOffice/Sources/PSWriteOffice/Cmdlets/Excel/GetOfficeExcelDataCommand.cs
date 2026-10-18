using System.Collections.Generic;
using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Reads worksheet data as dictionaries or PSCustomObjects.</summary>
/// <para>Uses the first row as headers and streams rows through the OfficeIMO Excel data reader.</para>
/// <example>
///   <summary>Read the used range as PSCustomObjects.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficeExcelData -Path .\report.xlsx -Sheet 'Summary' | Format-Table</code>
///   <para>Returns each row as a PSCustomObject with properties mapped from the header row.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeExcelData")]
public sealed class GetOfficeExcelDataCommand : PSCmdlet
{
    /// <summary>Workbook to read when it is already open.</summary>
    [Parameter(ValueFromPipeline = true)]
    public ExcelDocument? Document { get; set; }

    /// <summary>Path to the workbook when no <see cref="Document"/> is supplied.</summary>
    [Parameter(Position = 0)]
    public string? Path { get; set; }

    /// <summary>Worksheet name to read; defaults to the first sheet.</summary>
    [Parameter]
    public string? Sheet { get; set; }

    /// <summary>Optional A1 range (e.g. A1:D10). When omitted, the sheet's used range is read.</summary>
    [Parameter]
    public string? Range { get; set; }

    /// <summary>Prefer decimals (instead of doubles) for numeric cells.</summary>
    [Parameter]
    public SwitchParameter NumericAsDecimal { get; set; }

    /// <summary>Emit each row as a dictionary instead of PSCustomObjects.</summary>
    [Parameter]
    public SwitchParameter AsHashtable { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var document = Document;
        var ownsDocument = false;

        if (document == null)
        {
            if (string.IsNullOrWhiteSpace(Path))
            {
                throw new PSArgumentException("Specify -Path or provide an ExcelDocument on the pipeline.");
            }

            var resolved = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
            if (!File.Exists(resolved))
            {
                throw new FileNotFoundException($"File '{resolved}' was not found.", resolved);
            }

            document = ExcelDocument.Load(resolved, new ExcelLoadOptions
            {
                AccessMode = DocumentAccessMode.ReadOnly,
                PersistenceMode = DocumentPersistenceMode.Explicit
            });
            ownsDocument = true;
        }

        try
        {
            var options = new ExcelReadOptions
            {
                SheetName = string.IsNullOrWhiteSpace(Sheet) ? null : Sheet,
                SheetIndex = string.IsNullOrWhiteSpace(Sheet) ? 0 : null,
                A1Range = string.IsNullOrWhiteSpace(Range) ? null : Range,
                HasHeaderRow = true,
                NumericAsDecimal = NumericAsDecimal.IsPresent
            };

            var prevalidatedOutputProperties = false;
            using var reader = document.CreateDataReader(options);
            while (reader.Read())
            {
                var row = new Dictionary<string, object?>(reader.FieldCount, StringComparer.OrdinalIgnoreCase);
                for (var field = 0; field < reader.FieldCount; field++)
                {
                    var value = reader.GetValue(field);
                    row[reader.GetName(field)] = value == DBNull.Value ? null : value;
                }

                if (AsHashtable.IsPresent)
                {
                    WriteObject(row);
                }
                else
                {
                    var psObj = new PSObject();
                    foreach (KeyValuePair<string, object?> kv in row)
                    {
                        psObj.Properties.Add(new PSNoteProperty(kv.Key, kv.Value), prevalidatedOutputProperties);
                    }

                    prevalidatedOutputProperties = true;
                    WriteObject(psObj);
                }
            }
        }
        finally
        {
            if (ownsDocument)
            {
                document.Dispose();
            }
        }
    }
}
