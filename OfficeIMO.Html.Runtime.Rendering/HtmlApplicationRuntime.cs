using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime;

namespace OfficeIMO.Html.Runtime.Rendering;

/// <summary>Creates application runtime hosts with the providers shipped by the rendering package.</summary>
public static class HtmlApplicationRuntime {
    /// <summary>Creates the supported process runtime host over a deployed OfficeIMO worker.</summary>
    /// <remarks>
    /// The returned contract remains provider-neutral. This package currently wires the
    /// retained AngleSharp DOM services used to import frozen worker captures; callers do
    /// not need to select or reference that temporary provider directly.
    /// </remarks>
    public static IHtmlRuntimeHost CreateProcessHost(
        string workerAssemblyPath,
        string dotnetExecutable = "dotnet") =>
        new HtmlProcessRuntimeProvider(
            workerAssemblyPath,
            AngleSharpDomServices.Instance,
            dotnetExecutable);
}
