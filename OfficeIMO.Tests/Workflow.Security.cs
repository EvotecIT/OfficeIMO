using Xunit;

namespace OfficeIMO.Tests;

public class WorkflowSecurityTests {
    [Fact]
    public void VscodeExtensionWorkflow_UsesPinnedVsceForMarketplaceVersionCheck() {
        string workflow = ReadWorkflow();

        Assert.DoesNotContain("npm exec -- vsce", workflow, StringComparison.Ordinal);
        Assert.Contains("& node $VsceMain show $ExtensionId --json", workflow, StringComparison.Ordinal);
    }

    [Fact]
    public void VscodeExtensionWorkflow_ExposesVscePatOnlyToPublishStep() {
        string workflow = ReadWorkflow();
        string resolveStep = ExtractBetween(
            workflow,
            "      - name: Resolve marketplace publish target",
            "      - name: Publish VSIX");
        string publishStep = ExtractBetween(
            workflow,
            "      - name: Publish VSIX",
            "  attach-release-asset:");

        Assert.DoesNotContain("VSCE_PAT", resolveStep, StringComparison.Ordinal);
        Assert.Contains("VSCE_PAT: ${{ secrets.VSCE_PAT }}", publishStep, StringComparison.Ordinal);
        Assert.Equal(1, CountOccurrences(workflow, "VSCE_PAT: ${{ secrets.VSCE_PAT }}"));
    }

    private static string ReadWorkflow() =>
        File.ReadAllText(GetRepositoryPath(".github/workflows/vscode-extension.yml"));

    private static string ExtractBetween(string text, string start, string end) {
        int startIndex = text.IndexOf(start, StringComparison.Ordinal);
        Assert.True(startIndex >= 0, "Expected workflow section was not found: " + start);

        int endIndex = text.IndexOf(end, startIndex + start.Length, StringComparison.Ordinal);
        Assert.True(endIndex > startIndex, "Expected workflow section end was not found: " + end);

        return text.Substring(startIndex, endIndex - startIndex);
    }

    private static int CountOccurrences(string text, string value) {
        int count = 0;
        int index = 0;
        while ((index = text.IndexOf(value, index, StringComparison.Ordinal)) >= 0) {
            count++;
            index += value.Length;
        }

        return count;
    }

    private static string GetRepositoryPath(string relativePath) {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory != null) {
            string solutionPath = Path.Combine(directory.FullName, "OfficeIMO.sln");
            if (File.Exists(solutionPath)) {
                return Path.Combine(directory.FullName, relativePath.Replace('/', Path.DirectorySeparatorChar));
            }

            directory = directory.Parent;
        }

        throw new DirectoryNotFoundException("Unable to locate OfficeIMO repository root from test runtime base directory.");
    }
}
