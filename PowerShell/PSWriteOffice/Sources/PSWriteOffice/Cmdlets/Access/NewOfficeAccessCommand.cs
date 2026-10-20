using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Access;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Access;

/// <summary>Creates a native Access database model without a provider, seed file, or execution.</summary>
/// <example>
///   <summary>Creates a native Access database model without a provider, seed file, or execution.</summary>
///   <code>$database = New-OfficeAccess -Format Accdb -Title Contacts</code>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeAccess")]
[OutputType(typeof(AccessDocument))]
public sealed class NewOfficeAccessCommand : PSCmdlet {
    /// <summary>Native file family.</summary>
    [Parameter]
    public AccessFileFormat Format { get; set; } = AccessFileFormat.Accdb;

    /// <summary>Database title.</summary>
    [Parameter]
    public string? Title { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        WriteObject(AccessDocument.Create(new AccessCreateOptions { Format = Format, DatabaseTitle = Title }));
    }
}
