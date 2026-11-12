using System.Diagnostics;
using System.Globalization;
using System.Text;
using OfficeIMO.Yaml.Benchmarks.Adapters;
using OfficeIMO.Yaml.Tests.Conformance;

namespace OfficeIMO.Yaml.Benchmarks.Comparison;

/// <summary>
/// Checks that independent YAML readers accept what OfficeIMO.Yaml writes: new documents emitted from every valid suite
/// case and corpus file, and the outputs of the scripted edits. A reader counts for an input only when it loads the
/// input to the same values as OfficeIMO.Yaml; it then has to load OfficeIMO.Yaml's output to the values
/// OfficeIMO.Yaml reads from it.
/// </summary>
internal static class YamlIndependentReaders {
    private static readonly UTF8Encoding Utf8 = new(false);

    public sealed class Outcome {
        public Outcome(string reader, string output) {
            Reader = reader;
            Output = output;
        }

        public string Reader { get; }
        public string Output { get; }
        public int Applicable { get; set; }
        public int Agreed { get; set; }
        public List<string> Failures { get; } = new();
    }

    public static IReadOnlyList<Outcome> Run(YamlTestSuite suite, IReadOnlyList<YamlCorpusFile> corpus, string? python) {
        var officeImo = new OfficeIMOYamlConformanceEngine();
        IYamlConformanceEngine[] readers = { new YamlDotNetConformanceEngine(), new SharpYamlConformanceEngine() };
        var inputs = suite.Cases.Where(c => !c.IsError).Select(c => (Subject: c.Id, Bytes: c.Input))
            .Concat(corpus.Select(f => (Subject: f.Path, Bytes: f.Yaml))).ToList();

        var emitted = new List<(string Subject, byte[] Input, byte[] Output)>();
        var edited = new List<(string Subject, byte[] Input, byte[] Output)>();
        foreach ((string subject, byte[] bytes) in inputs) {
            try {
                emitted.Add((subject, bytes, officeImo.Emit(officeImo.Parse(bytes))));
            } catch (Exception) {
                // Inputs OfficeIMO.Yaml cannot read are covered by the conformance tables.
            }
            string text = YamlTextEncoding.Decode(bytes, out _);
            foreach (YamlEditResult result in YamlEditScripts.Run(subject, text, perKind: 2)) {
                if (result.Failure is null && result.Output is not null) edited.Add((subject + " " + result.Kind + " " + result.Target, bytes, Utf8.GetBytes(result.Output)));
            }
        }

        var outcomes = new List<Outcome>();
        foreach (IYamlConformanceEngine reader in readers) {
            outcomes.Add(Compare(reader, officeImo, emitted, "new documents"));
            outcomes.Add(Compare(reader, officeImo, edited, "edited documents"));
        }
        if (python != null) outcomes.AddRange(RunPython(python, officeImo, emitted, edited));
        return outcomes;
    }

    private static Outcome Compare(IYamlConformanceEngine reader, IYamlConformanceEngine officeImo, List<(string Subject, byte[] Input, byte[] Output)> outputs, string kind) {
        var outcome = new Outcome(reader.Name, kind);
        foreach ((string subject, byte[] input, byte[] output) in outputs) {
            // Only inputs both readers understand alike tell whether the other reader accepts our output.
            if (!Agrees(reader, officeImo, input)) continue;
            outcome.Applicable++;
            string? difference;
            try {
                difference = YamlJsonEquivalence.FindDifference(officeImo.Load(output), reader.Load(output));
            } catch (Exception exception) {
                difference = exception.GetType().Name + ": " + FirstLine(exception.Message);
            }
            if (difference == null) outcome.Agreed++;
            else outcome.Failures.Add(subject + ": " + difference);
        }
        return outcome;
    }

    private static bool Agrees(IYamlConformanceEngine reader, IYamlConformanceEngine officeImo, byte[] input) {
        try {
            return YamlJsonEquivalence.FindDifference(officeImo.Load(input), reader.Load(input)) == null;
        } catch (Exception) {
            return false;
        }
    }

    /// <summary>
    /// Loads every input and output with PyYAML (YAML 1.1) and ruamel.yaml (YAML 1.2) through the given interpreter and
    /// compares their JSON with OfficeIMO.Yaml's reading under the same YAML version, for the inputs where they agree.
    /// A missing module is reported as unavailable.
    /// </summary>
    private static IEnumerable<Outcome> RunPython(string python, IYamlConformanceEngine officeImo, List<(string Subject, byte[] Input, byte[] Output)> emitted, List<(string Subject, byte[] Input, byte[] Output)> edited) {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-yaml-python-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            foreach (string module in new[] { "yaml", "ruamel.yaml" }) {
                YamlSchema schema = module == "yaml" ? YamlSchema.Yaml11 : YamlSchema.Core;
                foreach ((string kind, var outputs) in new[] { ("new documents", emitted), ("edited documents", edited) }) {
                    var outcome = new Outcome(module == "yaml" ? "PyYAML" : "ruamel.yaml", kind);
                    for (int i = 0; i < outputs.Count; i++) {
                        File.WriteAllBytes(Path.Combine(directory, (2 * i).ToString(CultureInfo.InvariantCulture) + ".yaml"), outputs[i].Input);
                        File.WriteAllBytes(Path.Combine(directory, (2 * i + 1).ToString(CultureInfo.InvariantCulture) + ".yaml"), outputs[i].Output);
                    }
                    string? error = RunScript(python, module, directory, 2 * outputs.Count, out string[] results);
                    if (error != null) {
                        outcome.Failures.Add("unavailable: " + error);
                        yield return outcome;
                        continue;
                    }
                    for (int i = 0; i < outputs.Count; i++) {
                        string input = results[2 * i], output = results[2 * i + 1];
                        if (input.StartsWith("!", StringComparison.Ordinal) || !Same(Load(outputs[i].Input, schema), input)) continue;
                        outcome.Applicable++;
                        string? difference = output.StartsWith("!", StringComparison.Ordinal) ? "rejected the output" : Difference(Load(outputs[i].Output, schema), output);
                        if (difference == null) outcome.Agreed++;
                        else outcome.Failures.Add(outputs[i].Subject + ": " + difference);
                    }
                    yield return outcome;
                }
            }
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    private static IReadOnlyList<YamlConformanceNode>? Load(byte[] yaml, YamlSchema schema) {
        try {
            YamlDocument document = YamlDocument.Load(new MemoryStream(yaml, writable: false), new YamlLoadOptions { Schema = schema, DuplicateKeys = YamlDuplicateKeyPolicy.LastWins });
            return document.Documents.Select(part => OfficeIMOYamlConformanceEngine.ToConformance(part.Root)).ToList();
        } catch (Exception) {
            return null;
        }
    }

    private static bool Same(IReadOnlyList<YamlConformanceNode>? expected, string json) => Difference(expected, json) == null;

    private static string? Difference(IReadOnlyList<YamlConformanceNode>? expected, string json) {
        if (expected == null) return "OfficeIMO.Yaml did not load it";
        try {
            return YamlJsonEquivalence.FindDifference(expected, YamlJsonReader.ReadStream(json));
        } catch (Exception exception) {
            return exception.GetType().Name + ": " + FirstLine(exception.Message);
        }
    }

    private static string? RunScript(string python, string module, string directory, int count, out string[] results) {
        results = Array.Empty<string>();
        // YAML 1.2 core-schema documents with only JSON-representable values; an input the module rejects prints '!'.
        string loader = module == "yaml"
            ? "import yaml\ndef load(t): return list(yaml.safe_load_all(t))\n"
            : "from ruamel.yaml import YAML\ny = YAML(typ='safe', pure=True)\ndef load(t): return list(y.load_all(t))\n";
        string script = loader + "import json, sys, os, math, datetime\n"
            + "def norm(v):\n"
            + "  if isinstance(v, float) and (math.isnan(v) or math.isinf(v)): return str(v)\n"
            + "  if isinstance(v, (datetime.date, datetime.datetime)): return v.isoformat()\n"
            + "  if isinstance(v, dict): return {str(k): norm(x) for k, x in v.items()}\n"
            + "  if isinstance(v, list): return [norm(x) for x in v]\n"
            + "  return v\n"
            + "d = sys.argv[1]\n"
            + "for i in range(int(sys.argv[2])):\n"
            + "  try:\n"
            + "    docs = load(open(os.path.join(d, str(i) + '.yaml'), encoding='utf-8').read())\n"
            + "    print('\\n'.join(json.dumps(norm(x)) for x in docs).replace('\\n', '\\u2028'))\n"
            + "  except Exception:\n"
            + "    print('!')\n";
        string scriptPath = Path.Combine(directory, "load.py");
        File.WriteAllText(scriptPath, script, Utf8);
        var start = new ProcessStartInfo(python) { RedirectStandardOutput = true, RedirectStandardError = true, UseShellExecute = false, StandardOutputEncoding = Utf8 };
        start.ArgumentList.Add("-I");
        start.ArgumentList.Add(scriptPath);
        start.ArgumentList.Add(directory);
        start.ArgumentList.Add(count.ToString(CultureInfo.InvariantCulture));
        using Process process = Process.Start(start)!;
        string output = process.StandardOutput.ReadToEnd();
        string error = process.StandardError.ReadToEnd();
        process.WaitForExit();
        if (process.ExitCode != 0) return FirstLine(error);
        results = output.Split('\n').Take(count).Select(line => line.TrimEnd('\r').Replace("\u2028", "\n")).ToArray();
        return results.Length == count ? null : "the script printed " + results.Length + " results for " + count + " files";
    }

    private static string FirstLine(string text) {
        int newline = text.IndexOf('\n');
        return newline < 0 ? text : text.Substring(0, newline);
    }
}
