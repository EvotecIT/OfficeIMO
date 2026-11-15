using Xunit;

namespace OfficeIMO.Yaml.Tests;

public class YamlDocumentQueryTests {
    private const string Services = """
        defaults: &defaults
          port: 80
          enabled: true
        services:
          - name: web
            image: nginx:1.27
            <<: *defaults
            port: 8080
          - name: api
            image: api:2
            port: "8080"
            enabled: false
          - name: worker
            image: worker:1
            labels:
              "app.kubernetes.io/name": worker
        """;

    [Theory]
    [InlineData("services[*].name", "web,api,worker")]
    [InlineData("services.*.name", "web,api,worker")]
    [InlineData("$..image", "nginx:1.27,api:2,worker:1")]
    [InlineData("services[-1].name", "worker")]
    [InlineData("services[?(@.name == 'api')].image", "api:2")]
    [InlineData("services[?@.port == 8080].name", "web")]
    [InlineData("services[?(@.port == '8080')].name", "api")]
    [InlineData("services[?(@.port >= 100)].name", "web")]
    [InlineData("services[?(@.enabled == true)].name", "web")]
    [InlineData("services[?(!@.enabled)].name", "worker")]
    [InlineData("services[?(@.labels && @.name != 'web')].name", "worker")]
    [InlineData("services[?(@.name == 'web' || @.port == '8080')].name", "web,api")]
    [InlineData("services[?(@.labels[\"app.kubernetes.io/name\"] == 'worker')].image", "worker:1")]
    [InlineData("services[?(@.missing == null)].name", "")]
    [InlineData("services[?(@.name < 'b')].name", "api")]
    public void QueriesSelectNodesInDocumentOrder(string query, string expected) {
        IReadOnlyList<YamlQueryMatch> matches = YamlDocument.Parse(Services).Query(query);
        Assert.Equal(expected, string.Join(",", matches.Select(m => m.Node.AsString())));
    }

    [Fact]
    public void PathsAddressEachMatchAndNodesKeepTheirPositions() {
        YamlDocument document = YamlDocument.Parse(Services);

        IReadOnlyList<YamlQueryMatch> matches = document.Query("$..name");

        Assert.Equal(new[] { "services[0].name", "services[1].name", "services[2].name" }, matches.Select(m => m.Path));
        Assert.All(matches, m => Assert.Same(m.Node, document.Select(m.Path)));
        Assert.Equal(new[] { 5, 9, 13 }, matches.Select(m => m.Node.Start.Line));
        Assert.Equal("services[2].labels[\"app.kubernetes.io/name\"]", document.Query("..labels.*").Single().Path);
        Assert.Equal("$", document.Query("$").Single().Path);
    }

    [Fact]
    public void WildcardsSeeMergedKeysAndRecursiveDescentWalksTheNodesAsWritten() {
        YamlDocument document = YamlDocument.Parse(Services);

        // Own keys first, then merged keys the mapping does not override; the << key itself is not a value.
        Assert.Equal(new[] { "nginx:1.27", "8080", "web", "true" }.OrderBy(v => v),
            document.Query("services[0].*").Select(m => m.Node.AsString()).OrderBy(v => v));
        Assert.Equal("services[0].enabled", document.Query("services[?(@.enabled == true)].enabled").Single().Path);
        // The aliased defaults are visited once, where they are written.
        Assert.Equal(new[] { "defaults.port", "services[0].port", "services[1].port" }, document.Query("..port").Select(m => m.Path));
    }

    [Fact]
    public void EverySelectPathIsAQueryForTheSameNode() {
        YamlDocument document = YamlDocument.Parse(Services);
        foreach (string path in new[] { "services[0].image", "$.services[2].labels[\"app.kubernetes.io/name\"]", ".defaults.port", "services[0].enabled" }) {
            Assert.Same(document.Select(path), document.Query(path).Single().Node);
        }
        Assert.Empty(document.Query("services[7].name"));
        Assert.Empty(document.Query("defaults.port.deeper"));
    }

    [Fact]
    public void StreamQueriesPickDocumentsByIndexWildcardOrFilter() {
        YamlDocument document = YamlDocument.Parse("""
            kind: Deployment
            metadata: {name: web}
            ---
            kind: Service
            metadata: {name: web-svc}
            ---
            kind: Service
            metadata: {name: api-svc}
            """);

        IReadOnlyList<YamlQueryMatch> services = document.Query("$[?(@.kind == 'Service')].metadata.name");

        Assert.Equal(new[] { "$[1].metadata.name", "$[2].metadata.name" }, services.Select(m => m.Path));
        Assert.All(services, m => Assert.Same(m.Node, document.Select(m.Path)));
        Assert.Equal(new[] { "web", "web-svc", "api-svc" }, document.Query("$[*].metadata.name").Select(m => m.Node.AsString()));
        Assert.Equal("api-svc", document.Query("$[-1].metadata.name").Single().Node.AsString());
        Assert.Equal("Deployment", document.Query("kind").Single().Node.AsString());
    }

    [Fact]
    public void JsonDocumentsAnswerTheSameQueries() {
        YamlDocument document = YamlDocument.Parse("""
            {
              "services": [
                {"name": "web", "port": 8080},
                {"name": "api", "port": "8080"}
              ]
            }
            """);

        YamlQueryMatch match = document.Query("services[?(@.port == 8080)].name").Single();

        Assert.Equal("web", match.Node.AsString());
        Assert.Equal("services[0].name", match.Path);
        Assert.Equal(3, match.Node.Start.Line);
    }

    [Fact]
    public void NodeQueriesAreRelativeToTheNode() {
        YamlNode services = YamlDocument.Parse(Services).Select("services")!;

        YamlQueryMatch match = services.Query("[?(@.name == 'worker')].image").Single();

        Assert.Equal("[2].image", match.Path);
        Assert.Same(match.Node, services.Select(match.Path));
    }

    [Theory]
    [InlineData("services[")]
    [InlineData("services..")]
    [InlineData("services[?(@.name == )]")]
    [InlineData("services[?(@.name == 'web']")]
    [InlineData("services[?('web')]")]
    [InlineData("services[?($.name == 'web')]")]
    [InlineData("services['open]")]
    public void MalformedQueriesThrowWhateverTheData(string query) {
        Assert.Throws<ArgumentException>(() => YamlDocument.Parse(Services).Query(query));
        Assert.Throws<ArgumentException>(() => YamlDocument.Parse("{}").Query(query));
    }
}
