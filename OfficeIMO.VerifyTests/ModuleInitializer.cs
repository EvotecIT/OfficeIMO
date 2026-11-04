using System.Runtime.CompilerServices;
using VerifyXunit;

public static class ModuleInitializer
{
    [ModuleInitializer]
    public static void Init()
    {
        Verifier.UseProjectRelativeDirectory("verified");
    }
}