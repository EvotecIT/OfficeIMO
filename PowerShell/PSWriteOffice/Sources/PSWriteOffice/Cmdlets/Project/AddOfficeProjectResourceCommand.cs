using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Project;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Project;

/// <summary>Adds a work resource with an optional standard rate.</summary>
/// <example>
///   <summary>Adds a work resource with an optional standard rate.</summary>
///   <code>$project | Add-OfficeProjectResource -Name Engineer -StandardRate 125 -PassThru</code>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeProjectResource", DefaultParameterSetName = "Document")]
[OutputType(typeof(ProjectResource))]
public sealed class AddOfficeProjectResourceCommand : PSCmdlet {
    /// <summary>Project to inspect or change.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public ProjectDocument Document { get; set; } = null!;

    /// <summary>Resource name.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Name { get; set; } = string.Empty;

    /// <summary>Standard work rate in the project currency.</summary>
    [Parameter]
    [ValidateRange(0d, double.MaxValue)]
    public decimal? StandardRate { get; set; }

    /// <summary>Return the created or changed object.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var resource = Document.Resources.AddWork(Name);
        resource.StandardRate = StandardRate;
        resource.IsNull = false;
        Document.NormalizeDisplayIds();
        if (PassThru.IsPresent) WriteObject(resource);
    }
}
