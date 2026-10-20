using System.Reflection;
using System.Security.Cryptography;
using System.Xml.Linq;
using System.Text.RegularExpressions;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>Records the exact restored package and the notices shipped with that package.</summary>
internal static class PdfPeckerPackageEvidence {
    internal static HtmlPdfPackageEvidence Read() {
        string directory = Path.Combine(AppContext.BaseDirectory, "Licenses", "PdfPecker");
        XDocument specification = XDocument.Load(Path.Combine(directory, "pdfpecker.nuspec"));
        XElement metadata = specification.Root!.Elements().Single(element => element.Name.LocalName == "metadata");
        string Value(string name) => metadata.Elements().Single(element => element.Name.LocalName == name).Value;
        string version = Value("version");
        string expected = Assembly.GetExecutingAssembly().GetCustomAttributes<AssemblyMetadataAttribute>()
            .Single(attribute => attribute.Key == "PdfPeckerBenchmarkVersion").Value!;
        if (Value("id") != "PdfPecker" || version != expected) {
            throw new InvalidDataException("The PdfPecker package evidence does not match the compiled benchmark pin.");
        }
        string commit = metadata.Elements().Single(element => element.Name.LocalName == "repository").Attribute("commit")!.Value;
        string Hash(string file) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(Path.Combine(directory, file)))).ToLowerInvariant();
        Match license = Regex.Match(File.ReadAllText(Path.Combine(directory, "THIRD-PARTY-NOTICES.md")),
            "In SPDX terms the library is `([^`]+)`");
        if (!license.Success) throw new InvalidDataException("The restored PdfPecker notices do not declare their license expression.");
        return new("PdfPecker", version, commit, license.Groups[1].Value,
            Hash("LICENSE"), Hash("THIRD-PARTY-NOTICES.md"),
            File.ReadAllText(Path.Combine(directory, $"pdfpecker.{version}.nupkg.sha512")).Trim());
    }
}
