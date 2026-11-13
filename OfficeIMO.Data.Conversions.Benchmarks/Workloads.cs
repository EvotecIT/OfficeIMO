using System.Globalization;
using System.Text;

namespace OfficeIMO.Data.Conversions.Benchmarks;

/// <summary>Deterministic synthetic records: the shape of a typical API export, in every format the lanes read.</summary>
internal static class Workloads {
    public static string NestedJson(int count) {
        var builder = new StringBuilder(count * 260).Append('[');
        for (int i = 0; i < count; i++) {
            if (i > 0) builder.Append(',');
            builder.Append("{\"id\":").Append(i)
                .Append(",\"name\":\"Customer ").Append(i).Append('"')
                .Append(",\"email\":\"user").Append(i).Append("@example.com\"")
                .Append(",\"active\":").Append(i % 3 == 0 ? "false" : "true")
                .Append(",\"score\":").Append((i % 1000 / 10.0).ToString("0.0#", CultureInfo.InvariantCulture))
                .Append(",\"created\":\"2026-").Append((i % 12 + 1).ToString("D2", CultureInfo.InvariantCulture)).Append("-15T10:20:30Z\"")
                .Append(",\"address\":{\"city\":\"City ").Append(i % 97).Append("\",\"zip\":\"0").Append(i % 9000 + 1000).Append("\"}")
                .Append(",\"tags\":[\"a").Append(i % 5).Append("\",\"b").Append(i % 7).Append("\"]}");
        }

        return builder.Append(']').ToString();
    }

    public static string FlatJson(int count) {
        var builder = new StringBuilder(count * 150).Append('[');
        for (int i = 0; i < count; i++) {
            if (i > 0) builder.Append(',');
            builder.Append("{\"id\":").Append(i)
                .Append(",\"name\":\"Customer ").Append(i).Append('"')
                .Append(",\"email\":\"user").Append(i).Append("@example.com\"")
                .Append(",\"active\":").Append(i % 3 == 0 ? "false" : "true")
                .Append(",\"score\":").Append((i % 1000 / 10.0).ToString("0.0#", CultureInfo.InvariantCulture))
                .Append(",\"city\":\"City ").Append(i % 97).Append("\"}");
        }

        return builder.Append(']').ToString();
    }

    public static string Csv(int count) {
        var builder = new StringBuilder(count * 80).Append("id,name,email,active,score,city\n");
        for (int i = 0; i < count; i++) {
            builder.Append(i).Append(",Customer ").Append(i).Append(",user").Append(i).Append("@example.com,")
                .Append(i % 3 == 0 ? "false" : "true").Append(',')
                .Append((i % 1000 / 10.0).ToString("0.0#", CultureInfo.InvariantCulture)).Append(",City ").Append(i % 97).Append('\n');
        }

        return builder.ToString();
    }

    public static string Tsv(int count) => Csv(count).Replace(',', '\t');

    // A Cargo.lock-like file: one [[package]] table per record.
    public static string Toml(int count) {
        var builder = new StringBuilder(count * 260).Append("version = 4\n");
        for (int i = 0; i < count; i++) {
            builder.Append("\n[[package]]\nname = \"crate-").Append(i).Append("\"\nversion = \"1.").Append(i % 40).Append('.').Append(i % 7)
                .Append("\"\nsource = \"registry+https://github.com/rust-lang/crates.io-index\"\nchecksum = \"")
                .Append((i * 2654435761L).ToString("x16", CultureInfo.InvariantCulture)).Append("\"\ndownloads = ").Append(i * 37)
                .Append("\nscore = ").Append((i % 1000 / 10.0).ToString("0.0#", CultureInfo.InvariantCulture))
                .Append("\nyanked = ").Append(i % 11 == 0 ? "true" : "false")
                .Append("\npublished = 2026-").Append((i % 12 + 1).ToString("D2", CultureInfo.InvariantCulture)).Append("-15T10:20:30Z")
                .Append("\ndependencies = [\"crate-").Append(i % 5).Append("\", \"crate-").Append(i % 7).Append("\"]\n");
        }

        return builder.ToString();
    }

    // A Kubernetes-like multi-document stream.
    public static string Yaml(int documents) {
        var builder = new StringBuilder(documents * 400);
        for (int i = 0; i < documents; i++) {
            builder.Append("---\napiVersion: apps/v1\nkind: Deployment\nmetadata:\n  name: app-").Append(i)
                .Append("\n  labels:\n    app: web\n    tier: \"").Append(i % 3).Append("\"\nspec:\n  replicas: ").Append(i % 5 + 1)
                .Append("\n  template:\n    spec:\n      containers:\n        - name: web\n          image: nginx:1.").Append(i % 30)
                .Append("\n          ports:\n            - containerPort: 80\n          env:\n            - name: MODE\n              value: production\n            - name: RATIO\n              value: \"0.").Append(i % 10).Append("\"\n");
        }

        return builder.ToString();
    }

    public static string Xml(int count) {
        var builder = new StringBuilder(count * 220).Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>\n<customers>");
        for (int i = 0; i < count; i++) {
            builder.Append("<customer id=\"").Append(i).Append("\"><name>Customer ").Append(i).Append("</name><email>user").Append(i)
                .Append("@example.com</email><active>").Append(i % 3 == 0 ? "false" : "true").Append("</active><address><city>City ")
                .Append(i % 97).Append("</city><zip>0").Append(i % 9000 + 1000).Append("</zip></address><tag>a").Append(i % 5)
                .Append("</tag><tag>b").Append(i % 7).Append("</tag></customer>");
        }

        return builder.Append("</customers>").ToString();
    }
}
