using OfficeIMO.Core.Internal;
using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows;

public static partial class OfficeWorkflow {
    private static readonly Lazy<bool> ConverterRegistration = new(() => {
        foreach (OfficeWorkflowRoute route in OfficeWorkflowCatalog.ExecutableRoutes.Where(route => !OfficeConverter.IsBuiltinRoute(route.Id)))
            OfficeConverter.Register(new BuiltinWorkflowRoute(route.Id));
        return true;
    });

    /// <summary>Registers the additional Workflows conversion engines with <see cref="OfficeConverter"/>. Repeated calls reuse the same registrations.</summary>
    public static void RegisterConversions() => _ = ConverterRegistration.Value;

    internal static IOfficeConversionRoute CreateConversionRoute(OfficeWorkflowConversionRegistration registration) =>
        registration.DirectoryPackageInput == null ? new WorkflowRoute(registration.RouteId, registration)
            : new DirectoryWorkflowRoute(registration);

    private sealed class BuiltinWorkflowRoute(string id) : WorkflowRoute(id), IOfficeConversionCheckpointOwner {
        public string CaptureConfiguration(OfficeConversionOptions options) =>
            "FailOnLoss=" + options.FailOnLoss + "\n" + OfficeWorkflowCheckpointSettings.Capture(options, Id);

    }

    private sealed class DirectoryWorkflowRoute(OfficeWorkflowConversionRegistration registration)
        : WorkflowRoute(registration.RouteId, registration), IOfficeConversionDirectoryOwner {
        public OfficeConversionExecutionInput CaptureDirectoryInput(string path, OfficeConversionOptions options, OfficeOperationLimits limits) {
            OfficeWorkflowStreamInput input = Registration!.DirectoryPackageInput!(path, limits,
                Registration.SnapshotSettings(options.RegisteredSettings));
            return new OfficeConversionExecutionInput(input.Name, input.OpenRead, input.ExpectedSha256, input.SourcePublicationGuard);
        }
    }
}
