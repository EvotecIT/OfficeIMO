using System;
using System.Management.Automation;
using System.Security;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Services.Pdf;

/// <summary>Loads file inspection sources and leaves live documents under caller ownership.</summary>
internal static class PdfInspectionSource {
    internal static void Visit(PSCmdlet command, string source, string[] path, string[] literalPath,
        PdfDocument? document, PdfLoadOptions? options, SecureString? password, bool ignoreRestrictions, Action<PdfDocument> inspect) {
        if (source == "Document") {
            inspect(document ?? throw new PSArgumentException("Provide a live -Document."));
            return;
        }
        bool literal = source == "LiteralPath";
        var read = PdfCommandUtilities.CreateReadOptions(options, password, ignoreRestrictions);
        if (OfficeSwitch.Get(command.MyInvocation.BoundParameters, "IgnorePermissionRestrictions") is bool ignore)
            read = (read ?? new PdfLoadOptions()).Snapshot(null, null, ignore ? PdfPermissionPolicy.IgnoreRestrictions : PdfPermissionPolicy.Enforce);
        foreach (var file in OfficePath.ResolveInputs(command, literal ? literalPath : path, literal)) {
            using var owned = PdfDocument.Load(file, read);
            inspect(owned);
        }
    }
}
