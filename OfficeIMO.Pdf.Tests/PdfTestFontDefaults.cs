using System.Runtime.CompilerServices;
using OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

internal static class PdfTestFontDefaults {
    // Most generated-PDF tests pin standard-font metrics and resource names, so they must not depend on the
    // fonts installed on the test machine. Tests of default font embedding opt in with EmbedDefaultFonts = true.
#pragma warning disable CA2255 // The test assembly owns this start-up default.
    [ModuleInitializer]
    internal static void UseStandardFontsByDefault() => PdfCreateOptions.EmbedDefaultFontsInitially = false;
#pragma warning restore CA2255
}
