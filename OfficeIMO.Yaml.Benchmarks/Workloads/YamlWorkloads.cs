using System.Globalization;
using System.Text;
using OfficeIMO.Yaml.Tests.Conformance;

namespace OfficeIMO.Yaml.Benchmarks.Workloads;

/// <summary>Benchmark inputs. Real files come from the corpus; large inputs are generated deterministically at run time.</summary>
public enum YamlWorkload {
    /// <summary>One real GitHub Actions workflow (about 10 KB).</summary>
    SmallConfig,
    /// <summary>Generated multi-document Kubernetes stream of about 10 MB.</summary>
    KubernetesStream,
    /// <summary>Generated documents nesting 57 collections deep, under the 64-level default limit of SharpYaml's parser.</summary>
    DeepNesting,
    /// <summary>Generated literal and folded block scalars, about 2 MB.</summary>
    BlockScalars,
    /// <summary>Generated JSON written as YAML flow collections, about 2 MB.</summary>
    FlowHeavy,
    /// <summary>The real Airflow Compose file built on anchors, aliases and merge keys.</summary>
    AnchorHeavy,
    /// <summary>1,000 small real configuration files from the corpus, processed one by one.</summary>
    ManySmallFiles
}

internal static class YamlWorkloads {
    private static readonly Dictionary<YamlWorkload, string[]> Cache = new();

    public static string[] Get(YamlWorkload workload) {
        lock (Cache) {
            if (!Cache.TryGetValue(workload, out string[]? inputs)) {
                inputs = Create(workload);
                Cache.Add(workload, inputs);
            }

            return inputs;
        }
    }

    private static string[] Create(YamlWorkload workload) => workload switch {
        YamlWorkload.SmallConfig => new[] { CorpusText("GitHubActions/actions-checkout-test.yml") },
        YamlWorkload.KubernetesStream => new[] { KubernetesStream(10 * 1024 * 1024) },
        YamlWorkload.DeepNesting => new[] { DeepNesting(depth: 36, documents: 300) },
        YamlWorkload.BlockScalars => new[] { BlockScalars(2 * 1024 * 1024) },
        YamlWorkload.FlowHeavy => new[] { FlowHeavy(2 * 1024 * 1024) },
        YamlWorkload.AnchorHeavy => new[] { CorpusText("Compose/airflow-docker-compose.yaml") },
        YamlWorkload.ManySmallFiles => ManySmallFiles(1000),
        _ => throw new ArgumentOutOfRangeException(nameof(workload), workload, null)
    };

    private static string CorpusText(string path) {
        YamlCorpusFile file = YamlCorpus.Load().Single(f => f.Path == path);
        return new UTF8Encoding(false).GetString(file.Yaml);
    }

    // Small UTF-8 YAML files without a byte order mark, cycled until the count is reached.
    private static string[] ManySmallFiles(int count) {
        var sources = YamlCorpus.Load()
            .Where(f => f.DerivedFrom == null && f.Yaml.Length < 64 * 1024 && f.Yaml.Length > 0 && f.Yaml[0] != 0xEF && f.Yaml[0] != 0xFF)
            .Select(f => new UTF8Encoding(false).GetString(f.Yaml))
            .ToArray();
        var inputs = new string[count];
        for (int i = 0; i < count; i++) inputs[i] = sources[i % sources.Length];
        return inputs;
    }

    private static string KubernetesStream(int targetBytes) {
        var text = new StringBuilder(targetBytes + 4096);
        for (int i = 0; text.Length < targetBytes; i++) {
            string name = "service-" + i.ToString(CultureInfo.InvariantCulture);
            string tier = (i % 3) switch { 0 => "frontend", 1 => "backend", _ => "cache" };
            text.Append("---\napiVersion: apps/v1\nkind: Deployment\nmetadata:\n  name: ").Append(name)
                .Append("\n  namespace: team-").Append(i % 17)
                .Append("\n  labels:\n    app: ").Append(name).Append("\n    tier: ").Append(tier)
                .Append("\n  annotations:\n    deployment.kubernetes.io/revision: \"").Append(i % 9 + 1).Append("\"\n")
                .Append("spec:\n  replicas: ").Append(i % 5 + 1)
                .Append("\n  selector:\n    matchLabels:\n      app: ").Append(name)
                .Append("\n  template:\n    metadata:\n      labels:\n        app: ").Append(name).Append("\n        tier: ").Append(tier)
                .Append("\n    spec:\n      containers:\n      - name: ").Append(name)
                .Append("\n        image: registry.example.com/").Append(tier).Append(':').Append(i % 40).Append(".0.1")
                .Append("\n        args: [\"--port\", \"8080\", \"--verbose=").Append(i % 2 == 0 ? "true" : "false").Append("\"]")
                .Append("\n        env:\n        - name: GET_HOSTS_FROM\n          value: dns\n        - name: REPLICA_ID\n          value: '").Append(i).Append("'")
                .Append("\n        resources:\n          requests:\n            cpu: ").Append(100 + i % 400).Append("m\n            memory: ").Append(64 + i % 512).Append("Mi")
                .Append("\n        ports:\n        - containerPort: 8080\n          protocol: TCP\n")
                .Append("---\napiVersion: v1\nkind: Service\nmetadata:\n  name: ").Append(name)
                .Append("\n  labels: {app: ").Append(name).Append(", tier: ").Append(tier).Append("}\n")
                .Append("spec:\n  type: ClusterIP\n  ports:\n  - port: 80\n    targetPort: 8080\n  selector:\n    app: ").Append(name).Append('\n');
        }

        return text.ToString();
    }

    // Alternates a mapping key and a sequence entry holding a mapping, so each level adds one or two collections.
    private static string DeepNesting(int depth, int documents) {
        var text = new StringBuilder();
        for (int d = 0; d < documents; d++) {
            text.Append("---\n");
            int indent = 0;
            for (int level = 0; level < depth; level++) {
                if (level % 2 == 0) {
                    text.Append(' ', indent).Append("level").Append(level).Append(":\n");
                    indent += 2;
                } else {
                    text.Append(' ', indent).Append("- id: ").Append(d).Append('\n');
                    text.Append(' ', indent + 2).Append("next:\n");
                    indent += 4;
                }
            }

            text.Append(' ', indent).Append("leaf: [").Append(d).Append(", {deep: true}]\n");
        }

        return text.ToString();
    }

    private static string BlockScalars(int targetBytes) {
        var text = new StringBuilder(targetBytes + 8192);
        for (int i = 0; text.Length < targetBytes; i++) {
            bool literal = i % 2 == 0;
            text.Append("entry").Append(i).Append(literal ? ": |\n" : ": >-\n");
            for (int line = 0; line < 40; line++) {
                text.Append("  Line ").Append(line).Append(" of block ").Append(i)
                    .Append(literal ? " keeps its line break and \"quotes\" and: colons" : " folds into one paragraph with # no comment");
                if (line % 10 == 9) text.Append("\n");
                text.Append('\n');
            }
        }

        return text.ToString();
    }

    private static string FlowHeavy(int targetBytes) {
        var text = new StringBuilder(targetBytes + 1024);
        text.Append("[\n");
        for (int i = 0; text.Length < targetBytes; i++) {
            if (i > 0) text.Append(",\n");
            text.Append("  {\"id\": ").Append(i)
                .Append(", \"name\": \"item-").Append(i).Append("\", \"price\": ").Append((i * 1.25).ToString("0.00", CultureInfo.InvariantCulture))
                .Append(", \"active\": ").Append(i % 3 == 0 ? "false" : "true")
                .Append(", \"tags\": [\"t").Append(i % 7).Append("\", \"t").Append(i % 11).Append("\"], \"owner\": null")
                .Append(", \"dimensions\": {\"w\": ").Append(i % 100).Append(", \"h\": ").Append(i % 50).Append(", \"unit\": \"cm\"}}");
        }

        text.Append("\n]\n");
        return text.ToString();
    }
}
