using System.Globalization;
using System.Text;
using OfficeIMO.Yaml.Tests.Conformance;
using Xunit;

namespace OfficeIMO.Yaml.Tests;

// Paths are a grammar shared with OfficeIMO.Reader.Yaml; the suite has no notion of them.
public class YamlDocumentPathTests {
    private const string Deployment = """
        spec:
          template:
            spec:
              containers:
                - name: web
                  image: nginx:1.27
          selector: &labels
            "app.kubernetes.io/name": web
            "a\"b": quoted
            "line\nbreak": escaped
        extra: *labels
        """;

    [Theory]
    [InlineData("spec.template.spec.containers[0].image", "nginx:1.27")]
    [InlineData("$.spec.template.spec.containers[0].name", "web")]
    [InlineData(".spec.template.spec.containers[0].name", "web")]
    [InlineData("spec[\"template\"].spec.containers[0].name", "web")]
    [InlineData("spec.selector[\"app.kubernetes.io/name\"]", "web")]
    [InlineData("spec.selector[\"a\\\"b\"]", "quoted")]
    [InlineData("spec.selector[\"line\\nbreak\"]", "escaped")]
    [InlineData("spec.selector[\"line\\u000abreak\"]", "escaped")]
    [InlineData("extra[\"app.kubernetes.io/name\"]", "web")]
    public void PathsSelectNodes(string path, string expected) {
        Assert.Equal(expected, YamlDocument.Parse(Deployment).Select(path)!.AsString());
    }

    [Theory]
    [InlineData("spec.missing")]
    [InlineData("spec.template.spec.containers[1]")]
    [InlineData("spec.template.spec.containers.name")]
    [InlineData("spec[0]")]
    [InlineData("spec.template.spec.containers[0].image.deeper")]
    public void PathsThatMatchNothingReturnNull(string path) {
        Assert.Null(YamlDocument.Parse(Deployment).Select(path));
    }

    [Theory]
    [InlineData("missing..name")]
    [InlineData("missing.")]
    [InlineData("missing[x]")]
    [InlineData("missing[0")]
    [InlineData("missing[\"open")]
    [InlineData("missing[\"bad\\q\"]")]
    [InlineData("$name")]
    public void MalformedPathsThrowWhateverTheData(string path) {
        Assert.Throws<ArgumentException>(() => YamlDocument.Parse(Deployment).Select(path));
    }

    [Fact]
    public void SelectStartsAtTheNodeAndAliasesStayInTheTree() {
        YamlDocument document = YamlDocument.Parse(Deployment);
        YamlNode spec = document.Select("spec")!;
        Assert.Same(document.Select("spec.selector"), spec.Select("selector"));
        Assert.Same(spec, spec.Select("$"));
        Assert.IsType<YamlAlias>(document.Select("extra"));
    }

    [Fact]
    public void MultiDocumentPathsPickTheDocumentFirst() {
        YamlDocument document = YamlDocument.Parse("kind: A\n---\nkind: B\n---\n- kind: C\n");
        Assert.Equal("B", document.Select("$[1].kind")!.AsString());
        Assert.Equal("C", document.Select("$[2][0].kind")!.AsString());
        Assert.Equal("A", document.Select("kind")!.AsString());
        Assert.Null(document.Select("$[3].kind"));
        Assert.Equal("C", YamlDocument.Parse("- kind: C\n").Select("$[0].kind")!.AsString());
    }

    // Every node of every corpus file is addressed by the path OfficeIMO.Reader.Yaml would report for it, formatted with
    // the same rules as its TraverseYaml: "$" or "$[n]" per document, ".key" for identifier keys, ["key"] with JSON escapes
    // otherwise, and [i] for items. The Reader collapses whitespace runs in key text, so keys containing whitespace are
    // left out until the Reader runs on the engine.
    [Fact]
    public void ReaderPathsSelectTheSameNodesInEveryCorpusFile() {
        foreach (YamlCorpusFile file in YamlCorpus.Load()) {
            YamlDocument document = YamlDocument.Load(new MemoryStream(file.Yaml));
            for (int i = 0; i < document.Documents.Count; i++) {
                string root = document.Documents.Count > 1 ? "$[" + i.ToString(CultureInfo.InvariantCulture) + "]" : "$";
                CheckPaths(document, document.Documents[i].Root, root, file.Path);
            }
        }
    }

    private static void CheckPaths(YamlDocument document, YamlNode node, string path, string file) {
        Assert.True(ReferenceEquals(node, document.Select(path)), file + ": " + path);
        if (node is YamlSequence sequence) {
            for (int i = 0; i < sequence.Count; i++) CheckPaths(document, sequence[i], path + "[" + i.ToString(CultureInfo.InvariantCulture) + "]", file);
        } else if (node is YamlMapping mapping) {
            foreach (YamlMappingEntry entry in mapping) {
                if (entry.Key is not YamlScalar key) continue;
                string text = key.AsString();
                if (text.Any(char.IsWhiteSpace) || text.Length > 512) continue;
                CheckPaths(document, entry.Value, path + (IsIdentifier(text) ? "." + text : "[\"" + Escape(text) + "\"]"), file);
            }
        }
    }

    private static bool IsIdentifier(string text) =>
        text.Length > 0 && (text[0] == '_' || char.IsLetter(text[0])) && text.All(ch => ch == '_' || char.IsLetterOrDigit(ch));

    private static string Escape(string text) {
        var escaped = new StringBuilder();
        foreach (char ch in text) {
            switch (ch) {
                case '\\': escaped.Append("\\\\"); break;
                case '"': escaped.Append("\\\""); break;
                default:
                    if (char.IsControl(ch)) escaped.Append("\\u").Append(((int)ch).ToString("x4", CultureInfo.InvariantCulture));
                    else escaped.Append(ch);
                    break;
            }
        }
        return escaped.ToString();
    }
}
