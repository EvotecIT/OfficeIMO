using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO.Visio;

namespace OfficeIMO.Examples.Visio {
    public static class PremiumVisioShowcase {
        public static void Example_PremiumVisioShowcase(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Premium scenario showcase");
            string showcasePath = Path.Combine(folderPath, "Premium Visio Showcase");
            Directory.CreateDirectory(showcasePath);
            foreach (string filePath in Directory.EnumerateFiles(showcasePath, "*.vsdx", SearchOption.TopDirectoryOnly)) {
                File.Delete(filePath);
            }

            IReadOnlyList<VisioGalleryResult> results = VisioPremiumGallery.Create(showcasePath);
            ValidateGeneratedPackages(showcasePath, results);

            string? firstFile = results.Select(result => result.FilePath).FirstOrDefault();
            if (openVisio && firstFile != null) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(firstFile) { UseShellExecute = true });
            }
        }

        private static void ValidateGeneratedPackages(string showcasePath, IReadOnlyList<VisioGalleryResult> results) {
            string summaryPath = Path.Combine(showcasePath, "premium-showcase-summary.md");
            using StreamWriter writer = new(summaryPath, false);
            writer.WriteLine("# Premium Visio Showcase Summary");
            writer.WriteLine();
            writer.WriteLine($"Generated: {DateTimeOffset.UtcNow:O}");
            writer.WriteLine($"VSDX files: {results.Count}");
            writer.WriteLine();
            writer.WriteLine("## Packages");
            writer.WriteLine();

            foreach (VisioGalleryResult result in results.OrderBy(result => result.FilePath, StringComparer.OrdinalIgnoreCase)) {
                if (!result.IsClean) {
                    IEnumerable<string> issues = result.PackageIssues
                        .Concat(result.QualityIssues.Select(issue => issue.ToString()));
                    string message = string.Join(Environment.NewLine, issues.Select(issue => "      " + issue));
                    throw new InvalidOperationException($"Premium Visio example failed validation: {result.FilePath}{Environment.NewLine}{message}");
                }

                writer.WriteLine($"- `{Path.GetRelativePath(showcasePath, result.FilePath)}`");
                Console.WriteLine($"    premium package ok: {Path.GetFileName(result.FilePath)}");
            }

            Console.WriteLine($"    premium summary: {summaryPath}");
        }
    }
}
