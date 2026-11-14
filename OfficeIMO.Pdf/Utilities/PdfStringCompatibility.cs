#if NETFRAMEWORK || NETSTANDARD2_0
namespace OfficeIMO.Pdf;

// .NET Framework strings have no Contains(char), so text.Contains('x') bound to the LINQ enumerator, which is much
// slower than IndexOf. This overload is closer in scope than System.Linq and takes precedence.
internal static class PdfStringCompatibility {
    internal static bool Contains(this string value, char character) => value.IndexOf(character) >= 0;
}
#endif
