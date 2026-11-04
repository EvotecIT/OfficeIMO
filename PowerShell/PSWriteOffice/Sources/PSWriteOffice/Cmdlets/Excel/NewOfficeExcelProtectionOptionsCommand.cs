using System.Management.Automation;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Creates Excel-specific hash and reservation settings for Protect-OfficeDocument.</summary>
/// <example><summary>Set the name displayed by Excel for a protected write reservation.</summary>
/// <code>$options = New-OfficeExcelProtectionOptions -ReservedBy 'Finance team'
/// Protect-OfficeDocument -Path ./Budget.xlsx -Restriction ReadOnlyRecommended -Options $options -OutputPath ./Budget-review.xlsx</code></example>
[Cmdlet(VerbsCommon.New, "OfficeExcelProtectionOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(OfficeProtectionExcelOptions))]
public sealed class NewOfficeExcelProtectionOptionsCommand : OfficeCmdlet {
    /// <summary>Precomputed native workbook protection hash, with at most four hexadecimal digits.</summary>
    [Parameter(ParameterSetName = "Default")]
    public string? LegacyPasswordHash { get; set; }
    /// <summary>Precomputed native write-reservation hash, with at most four hexadecimal digits.</summary>
    [Parameter]
    public string? ModifyPasswordHash { get; set; }
    /// <summary>Name displayed by Excel for the write reservation.</summary>
    [Parameter]
    public string? ReservedBy { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        OfficeProtectionOptionBinding.ValidateHash(LegacyPasswordHash);
        OfficeProtectionOptionBinding.ValidateHash(ModifyPasswordHash);
        WriteObject(new OfficeProtectionExcelOptions {
            LegacyPasswordHash = LegacyPasswordHash, ModifyPasswordHash = ModifyPasswordHash, ReservedBy = ReservedBy
        });
    }
}
