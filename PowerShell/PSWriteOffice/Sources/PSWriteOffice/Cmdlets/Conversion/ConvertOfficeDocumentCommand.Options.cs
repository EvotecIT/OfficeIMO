using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Net;
using OfficeIMO;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Conversion;

public sealed partial class ConvertOfficeDocumentCommand {
    private OfficeConversionOptions CaptureOptions(IReadOnlyList<string> inputs, string target,
        IReadOnlyList<OfficeConversionCapability>? selectedRoutes = null) {
        var inputFormat = CaptureInputFormat();
        bool hasDirectory = inputs.Any(Directory.Exists);
        var routes = selectedRoutes ?? OfficeConverter.Routes.Where(route =>
            (Route == null || string.Equals(route.Id, Route, StringComparison.OrdinalIgnoreCase)) &&
            string.Equals(route.TargetExtension, target, StringComparison.OrdinalIgnoreCase) &&
            (hasDirectory ? Extension == null || route.SourceExtensions.Intersect(Extension.Select(NormalizeExtension), StringComparer.OrdinalIgnoreCase).Any()
                : inputs.Any(input => route.SourceExtensions.Contains(System.IO.Path.GetExtension(OfficeConverter.ConversionSourceName(input, inputFormat)), StringComparer.OrdinalIgnoreCase))))
            .ToArray();
        if (Route != null && routes.Count == 0) throw new PSArgumentException("The selected route does not match the source and target extensions.");
        var engine = (EngineIntrinsics)SessionState.PSVariable.GetValue("ExecutionContext");
        var result = OfficeConversionOptionBinding.Capture(engine, Options, routes);
        if (From != null) { result.SourceFormat = inputFormat.SourceFormat; result.SourceExtension = inputFormat.SourceExtension; }
        if (MyInvocation.BoundParameters.ContainsKey(nameof(Password)))
            result.SourcePassword = Password == null ? null : new NetworkCredential(string.Empty, Password).Password;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(FailOnLoss))) result.FailOnLoss = FailOnLoss.IsPresent;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(CompressPdf))) result.CompressPdfOutput = CompressPdf.IsPresent;
        result.WordLoad ??= new OfficeIMO.Word.WordLoadOptions { FailOnLoss = true };
        if (MyInvocation.BoundParameters.ContainsKey(nameof(AllowLegacyImportLoss))) result.WordLoad.FailOnLoss = !AllowLegacyImportLoss.IsPresent;
        if (Route != null) result.RouteId = Route;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(OutputProfile))) result.OutputProfile = OutputProfile;
        result.FileConflictPolicy = ConflictPolicy;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(MaximumInputBytes))) result.MaximumInputBytes = MaximumInputBytes;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(MaximumOutputBytes))) result.MaximumOutputBytes = MaximumOutputBytes;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(MaximumXmlCharactersInPart))) result.MaximumXmlCharactersInPart = MaximumXmlCharactersInPart;
        return result;
    }

    private OfficeConversionOptions CaptureInputFormat() {
        object? value = Options is { Length: 1 } ? Options[0] : null;
        while (value is PSObject wrapped && !ReferenceEquals(value, wrapped.BaseObject)) value = wrapped.BaseObject;
        var result = value is OfficeConversionOptions options ? options.Clone() : new OfficeConversionOptions();
        if (OfficeFormatSelection.Parse(From) is { } from) {
            result.SourceFormat = from.Family;
            result.SourceExtension = from.PhysicalExtension;
        }
        return result;
    }

    private static string NormalizeExtension(string extension) => extension.StartsWith(".", StringComparison.Ordinal) ? extension : "." + extension;
}
