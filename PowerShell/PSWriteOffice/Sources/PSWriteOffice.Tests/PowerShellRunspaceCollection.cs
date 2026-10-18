namespace PSWriteOffice.Tests;

[CollectionDefinition(Name, DisableParallelization = true)]
public sealed class PowerShellRunspaceCollection
{
    public const string Name = "PowerShell runspace";
}
