using System;
using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.VerifyTests;
using VerifyTests;
using VerifyXunit;

public static class ModuleInitializer
{
    [ModuleInitializer]
    public static void Init()
    {
        // To disable Visual Studio popping up on every test execution.
        var diffDisabled = Environment.GetEnvironmentVariable("DiffEngine_Disabled");
        if (string.IsNullOrEmpty(diffDisabled))
            Environment.SetEnvironmentVariable("DiffEngine_Disabled", "true");

        Environment.SetEnvironmentVariable("Verify_DisableClipboard", "true");

        Verifier.UseSourceFileRelativeDirectory("verified");
        VerifierSettings.RegisterFileConverter<WordprocessingDocument>(WordprocessingDocumentConverter.Convert);
    }
}
