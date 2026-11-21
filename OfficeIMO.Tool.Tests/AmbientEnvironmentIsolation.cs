using System.Runtime.CompilerServices;
using OfficeIMO.Tool.Agent;

namespace OfficeIMO.Tool.Tests;

/// <summary>
/// Keeps the tool's ambient configuration on the developer's machine out of the tests. Commands run in process
/// and read <c>OFFICEIMO_MCP_ALLOWED_ROOTS</c>, so a value configured for a local MCP client would otherwise
/// refuse every test file under the temporary directory. Tests that need roots pass them explicitly, either to
/// <see cref="AgentPathPolicy"/> or to the child process environment.
/// </summary>
internal static class AmbientEnvironmentIsolation {
    [ModuleInitializer]
    internal static void ClearAmbientToolConfiguration() =>
        Environment.SetEnvironmentVariable(AgentPathPolicy.AllowedRootsEnvironmentVariable, null);
}
