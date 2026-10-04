using System.Runtime.CompilerServices;

namespace OfficeIMO.Email.Tests;

internal static class EmailTestRepository {
    // The artifact output root can live outside the checkout; the compiler's source path still locates fixtures.
    internal static string FindRoot([CallerFilePath] string sourceFile = "") {
        foreach (string? candidate in new[] { AppContext.BaseDirectory, Path.GetDirectoryName(sourceFile), Directory.GetCurrentDirectory() }) {
            DirectoryInfo? directory = candidate == null ? null : new DirectoryInfo(candidate);
            while (directory != null) {
                if (File.Exists(Path.Combine(directory.FullName, "OfficeIMO.sln"))) return directory.FullName;
                directory = directory.Parent;
            }
        }
        throw new DirectoryNotFoundException("The OfficeIMO source checkout required by this test was not found.");
    }
}
