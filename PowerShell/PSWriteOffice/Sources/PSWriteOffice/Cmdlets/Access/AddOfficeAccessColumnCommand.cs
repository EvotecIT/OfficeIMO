using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Access;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Access;

/// <summary>Adds a typed field before rows are appended.</summary>
/// <example>
///   <summary>Adds a typed field before rows are appended.</summary>
///   <code>$table | Add-OfficeAccessColumn -Name DisplayName -Type ShortText -MaxLength 120</code>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeAccessColumn", DefaultParameterSetName = "Table")]
[OutputType(typeof(AccessColumn))]
[Alias("AccessColumn") ]
public sealed class AddOfficeAccessColumnCommand : OfficeCmdlet {
    /// <summary>Table in a newly created database model.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = "Table")]
    public AccessTable? Table { get; set; }

    /// <summary>Field name.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Name { get; set; } = string.Empty;

    /// <summary>Native field type.</summary>
    [Parameter]
    public AccessDataType Type { get; set; } = AccessDataType.ShortText;

    /// <summary>Maximum short-text length.</summary>
    [Parameter]
    [ValidateRange(1, 255)]
    public int? MaxLength { get; set; }

    /// <summary>First sequential AutoNumber value.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int AutoNumberSeed { get; set; } = 1;

    /// <summary>Decimal precision.</summary>
    [Parameter]
    [ValidateRange(1, 28)]
    public int Precision { get; set; } = 28;

    /// <summary>Decimal scale.</summary>
    [Parameter]
    [ValidateRange(0, 28)]
    public int Scale { get; set; }

    /// <summary>Return the created or changed object.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var table = ResolveDocument(Table);
        if (MyInvocation.BoundParameters.ContainsKey(nameof(MaxLength)) && Type != AccessDataType.ShortText)
            throw new PSArgumentException("MaxLength applies only to ShortText columns.");
        if (MyInvocation.BoundParameters.ContainsKey(nameof(AutoNumberSeed)) && Type != AccessDataType.AutoNumber)
            throw new PSArgumentException("AutoNumberSeed applies only to AutoNumber columns.");
        if ((MyInvocation.BoundParameters.ContainsKey(nameof(Precision)) || MyInvocation.BoundParameters.ContainsKey(nameof(Scale))) && Type != AccessDataType.Decimal)
            throw new PSArgumentException("Precision and Scale apply only to Decimal columns.");
        if (Type == AccessDataType.Decimal && Scale > Precision)
            throw new PSArgumentException("Scale cannot exceed Precision.");
        var column = Type == AccessDataType.AutoNumber ? table.Columns.AddAutoNumber(Name, AutoNumberSeed)
            : Type == AccessDataType.Decimal ? table.Columns.AddDecimal(Name, Precision, Scale)
            : table.Columns.Add(Name, Type, MaxLength);
        if (PassThru.IsPresent) WriteObject(column);
    }
}
