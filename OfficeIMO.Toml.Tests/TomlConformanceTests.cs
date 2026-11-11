using OfficeIMO.Toml.Tests.Conformance;
using Xunit;
using Xunit.Abstractions;

namespace OfficeIMO.Toml.Tests;

// toml-test (TOML 1.1.0) is the oracle for reading. Each test runs every case and compares the failures with the
// checked-in exception list, so a regression fails and a fixed case must be removed from the list.
public class TomlConformanceTests {
    private readonly ITestOutputHelper _output;

    public TomlConformanceTests(ITestOutputHelper output) {
        _output = output;
    }

    [Fact]
    public void ValidCasesDecodeToTheExpectedJson() {
        var failures = new SortedDictionary<string, string>(StringComparer.Ordinal);
        foreach (TomlTestCase testCase in TomlTestSuite.Valid) {
            try {
                using TomlDocument document = TomlDocument.Load(new MemoryStream(testCase.Toml));
                string? difference = TomlTestJson.Compare(document.Root, JsonText.Parse(testCase.Json!));
                if (difference != null) failures[testCase.Name] = difference;
            } catch (Exception ex) {
                failures[testCase.Name] = ex.GetType().Name + ": " + ex.Message;
            }
        }
        Report("valid", TomlTestSuite.Valid.Count, failures);
    }

    [Fact]
    public void InvalidCasesAreRejectedWithTomlException() {
        var failures = new SortedDictionary<string, string>(StringComparer.Ordinal);
        foreach (TomlTestCase testCase in TomlTestSuite.Invalid) {
            try {
                TomlDocument.Load(new MemoryStream(testCase.Toml));
                failures[testCase.Name] = "accepted";
            } catch (TomlException) {
            } catch (Exception ex) {
                failures[testCase.Name] = "threw " + ex.GetType().Name + " instead of TomlException: " + ex.Message;
            }
        }
        Report("invalid", TomlTestSuite.Invalid.Count, failures);
    }

    [Fact]
    public void ValidCasesSaveBackByteForByte() {
        var failures = new SortedDictionary<string, string>(StringComparer.Ordinal);
        foreach (TomlTestCase testCase in TomlTestSuite.Valid) {
            string? difference = RoundTrip(testCase.Toml);
            if (difference != null) failures[testCase.Name] = difference;
        }
        Report("round-trip", TomlTestSuite.Valid.Count, failures);
    }

    [Fact]
    public void CorpusFilesSaveBackByteForByte() {
        IReadOnlyList<TomlCorpusFile> corpus = TomlCorpus.Load();
        Assert.NotEmpty(corpus);
        foreach (TomlCorpusFile file in corpus) Assert.True(RoundTrip(file.Bytes) is null, file.Path + ": " + RoundTrip(file.Bytes));
    }

    internal static string? RoundTrip(byte[] bytes) {
        try {
            using TomlDocument document = TomlDocument.Load(new MemoryStream(bytes));
            var output = new MemoryStream();
            document.Save(output);
            return output.ToArray().AsSpan().SequenceEqual(bytes) ? null : "saved bytes differ";
        } catch (Exception ex) {
            return ex.GetType().Name + ": " + ex.Message;
        }
    }

    private void Report(string kind, int total, SortedDictionary<string, string> failures) {
        HashSet<string> known = TomlTestSuite.ReadList("known-failures.txt");
        _output.WriteLine(kind + ": " + (total - failures.Count) + "/" + total + " pass");
        var unexpected = failures.Where(f => !known.Contains(f.Key)).Select(f => f.Key + " -> " + f.Value).ToList();
        var fixedCases = known.Where(name => name.StartsWith(kind == "round-trip" ? "valid/" : kind + "/", StringComparison.Ordinal) && !failures.ContainsKey(name)).ToList();
        foreach (string line in unexpected) _output.WriteLine("FAIL " + line);
        Assert.True(unexpected.Count == 0, unexpected.Count + " unexpected " + kind + " failures:\n" + string.Join("\n", unexpected.Take(40)));
        Assert.True(fixedCases.Count == 0 || kind == "round-trip", "Listed cases now pass; remove them from known-failures.txt:\n" + string.Join("\n", fixedCases));
    }
}
