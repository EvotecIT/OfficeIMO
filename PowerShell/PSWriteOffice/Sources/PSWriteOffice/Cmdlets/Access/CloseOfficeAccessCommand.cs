using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Access;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Access;

/// <summary>Closes a native database without implicitly saving it.</summary>
/// <example>
///   <summary>Closes a native database without implicitly saving it.</summary>
///   <code>$database | Close-OfficeAccess</code>
/// </example>
[Cmdlet(VerbsCommon.Close, "OfficeAccess", DefaultParameterSetName = "Document")]
public sealed class CloseOfficeAccessCommand : PSCmdlet {
    /// <summary>Live database document.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public AccessDocument Document { get; set; } = null!;

    /// <inheritdoc />
    protected override void ProcessRecord() {
        Document.Dispose();
    }
}
