using System;
using System.IO;
using System.Runtime.CompilerServices;

namespace OfficeIMO.Tests;

// Source-based tests can run from an external MSBuild artifacts directory.
internal static class RepositoryTestPaths {
    internal static string Find([CallerFilePath] string sourceFile = "") {
        string? configured = Environment.GetEnvironmentVariable("OFFICEIMO_TEST_REPOSITORY_ROOT");
        if (!string.IsNullOrWhiteSpace(configured)) {
            if (!IsRepository(configured)) throw new DirectoryNotFoundException("OFFICEIMO_TEST_REPOSITORY_ROOT is not an OfficeIMO source checkout.");
            return Path.GetFullPath(configured);
        }

        foreach (string start in new[] { Path.GetDirectoryName(sourceFile) ?? "", Directory.GetCurrentDirectory(), AppContext.BaseDirectory }) {
            if (string.IsNullOrWhiteSpace(start) || !Directory.Exists(start)) continue;
            for (DirectoryInfo? directory = new DirectoryInfo(start); directory != null; directory = directory.Parent) {
                if (IsRepository(directory.FullName)) return directory.FullName;
            }
        }

        throw new DirectoryNotFoundException("Could not locate the OfficeIMO source checkout; set OFFICEIMO_TEST_REPOSITORY_ROOT when running copied test assemblies.");
    }

    private static bool IsRepository(string path) =>
        File.Exists(Path.Combine(path, "OfficeIMO.sln")) &&
        File.Exists(Path.Combine(path, "OfficeIMO.Core", "OfficeIMO.Core.csproj"));
}
