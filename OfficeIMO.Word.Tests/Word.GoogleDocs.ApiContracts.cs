using System.Reflection;
using OfficeIMO.Word.GoogleDocs;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void GoogleDocsPlanningApisUseBuildVocabulary() {
        string[] names = typeof(WordGoogleDocsExtensions)
            .GetMethods(BindingFlags.Public | BindingFlags.Static | BindingFlags.DeclaredOnly)
            .Select(static method => method.Name)
            .ToArray();

        Assert.Contains("BuildGoogleDocsPlan", names);
        Assert.Contains("BuildGoogleDocsBatch", names);
        Assert.DoesNotContain("CreateGoogleDocsTranslationPlan", names);
        Assert.DoesNotContain("CreateGoogleDocsBatch", names);
    }
}
