using System.Runtime.CompilerServices;

namespace OfficeIMO.Workflows.Tests;

internal static class ConversionRegistration {
    [ModuleInitializer]
    internal static void Register() => OfficeWorkflow.RegisterConversions();
}
