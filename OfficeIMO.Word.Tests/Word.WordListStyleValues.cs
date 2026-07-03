using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void WordListStyle_IntegerValues() {
        Assert.Equal(12, (int)WordListStyle.Custom);
        Assert.Equal(13, (int)WordListStyle.Numbered);
    }
}

