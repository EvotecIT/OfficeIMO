using System.Text;
using Xunit;

namespace OfficeIMO.Yaml.Tests;

// The suite contains no alias bomb, cycle or oversized input, so the model's own limits are proven here.
public class YamlDocumentLimitTests {
    /// <summary>Nine levels of nine aliases each: about 387 million nodes when resolved, a few hundred bytes of source.</summary>
    private static string BillionLaughs() {
        var text = new StringBuilder("a: &a [\"lol\",\"lol\",\"lol\",\"lol\",\"lol\",\"lol\",\"lol\",\"lol\",\"lol\"]\n");
        for (char level = 'b'; level <= 'i'; level++) {
            char previous = (char)(level - 1);
            text.Append(level).Append(": &").Append(level).Append(" [");
            for (int i = 0; i < 9; i++) text.Append(i == 0 ? "*" : ",*").Append(previous);
            text.Append("]\n");
        }
        return text.ToString();
    }

    [Fact]
    public void AliasBombFailsWhileLoadingWithoutExpanding() {
        YamlLimitExceededException error = Assert.Throws<YamlLimitExceededException>(() => YamlDocument.Parse(BillionLaughs()));
        Assert.Equal(nameof(YamlLoadOptions.MaxAliasExpansion), error.LimitName);
        Assert.Equal(100_000, error.Limit);
        Assert.True(error.Line > 1);
    }

    // Each alias counts the nodes its anchor stands for, nested aliases included: *a stands for 2 nodes and *b for
    // 1 + 2 + 2 = 5, so the three aliases stand for 9.
    [Fact]
    public void AliasExpansionCountsNestedAliasesAtEveryUse() {
        const string yaml = "a: &a [x]\nb: &b [*a, *a]\nc: *b\n";
        YamlDocument.Parse(yaml, new YamlLoadOptions { MaxAliasExpansion = 9 });
        Assert.Throws<YamlLimitExceededException>(() => YamlDocument.Parse(yaml, new YamlLoadOptions { MaxAliasExpansion = 8 }));
    }

    [Fact]
    public void AliasExpansionIsCountedPerDocument() {
        const string document = "a: &a [x, y]\nb: [*a, *a, *a]\n";
        YamlDocument.Parse(document + "---\n" + document, new YamlLoadOptions { MaxAliasExpansion = 9 });
    }

    [Fact]
    public void AliasInsideItsOwnAnchorIsRejectedAsInfiniteExpansion() {
        YamlLimitExceededException error = Assert.Throws<YamlLimitExceededException>(() => YamlDocument.Parse("&a [x, *a]\n"));
        Assert.Equal(nameof(YamlLoadOptions.MaxAliasExpansion), error.LimitName);
    }

    [Theory]
    [InlineData("a: *missing\n")]
    [InlineData("a: &x 1\n---\nb: *x\n")]
    public void AliasWithoutAnAnchorBeforeItInItsDocumentFails(string yaml) {
        YamlException error = Assert.Throws<YamlException>(() => YamlDocument.Parse(yaml));
        Assert.Equal(YamlErrorCodes.UndefinedAlias, error.Code);
    }

    // "a: [1, 2]" has five nodes: the mapping, the key, the sequence and two items. Aliases count as nodes, and each
    // document of a stream has its own count.
    [Theory]
    [InlineData("a: [1, 2]\n", 5)]
    [InlineData("a: &x 1\nb: *x\n", 5)]
    [InlineData("a: [1, 2]\n---\na: [1, 2]\n", 5)]
    public void NodeCountBeyondTheLimitFails(string yaml, int nodes) {
        YamlDocument.Parse(yaml, new YamlLoadOptions { MaxNodes = nodes });
        YamlLimitExceededException error = Assert.Throws<YamlLimitExceededException>(() => YamlDocument.Parse(yaml, new YamlLoadOptions { MaxNodes = nodes - 1 }));
        Assert.Equal(nameof(YamlLoadOptions.MaxNodes), error.LimitName);
    }

    [Fact]
    public async Task StreamLargerThanTheInputLimitFailsForSyncAndAsyncLoads() {
        byte[] yaml = Encoding.UTF8.GetBytes("a: " + new string('b', 200) + "\n");
        var options = new YamlLoadOptions { MaxInputBytes = 100 };
        Assert.Throws<YamlLimitExceededException>(() => YamlDocument.Load(new MemoryStream(yaml), options));
        await Assert.ThrowsAsync<YamlLimitExceededException>(() => YamlDocument.LoadAsync(new MemoryStream(yaml), options));
    }
}
