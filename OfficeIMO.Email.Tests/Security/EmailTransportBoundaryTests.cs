namespace OfficeIMO.Email.Tests;

public sealed class EmailTransportBoundaryTests {
    [Fact]
    public void ArtifactEngineReferencesSecurityButNoMailOrCloudTransportSdk() {
        string root = FindRepositoryRoot();
        string project = File.ReadAllText(Path.Combine(root, "OfficeIMO.Email", "OfficeIMO.Email.csproj"));
        Assert.Contains("OfficeIMO.Security.csproj", project, StringComparison.Ordinal);
        Assert.DoesNotContain("MailKit", project, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("MimeKit", project, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Microsoft.Graph", project, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Google.Apis", project, StringComparison.OrdinalIgnoreCase);
        string googleCore = File.ReadAllText(Path.Combine(root, "OfficeIMO.GoogleWorkspace", "OfficeIMO.GoogleWorkspace.csproj"));
        Assert.DoesNotContain("Google.Apis", googleCore, StringComparison.OrdinalIgnoreCase);
    }

    private static string FindRepositoryRoot() {
        DirectoryInfo? directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory != null && !File.Exists(Path.Combine(directory.FullName, "OfficeIMO.sln"))) directory = directory.Parent;
        return directory?.FullName ?? throw new DirectoryNotFoundException("OfficeIMO repository root was not found.");
    }
}
