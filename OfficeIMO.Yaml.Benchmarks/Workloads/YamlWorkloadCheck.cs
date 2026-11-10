namespace OfficeIMO.Yaml.Benchmarks.Workloads;

/// <summary>Runs every benchmark validation once without timing: <c>--validate-workloads</c>.</summary>
internal static class YamlWorkloadCheck {
    public static int Run() {
        int failures = 0;
        foreach (YamlWorkload workload in Enum.GetValues(typeof(YamlWorkload))) {
            string[] inputs = YamlWorkloads.Get(workload);
            long bytes = inputs.Distinct(StringComparer.Ordinal).Sum(input => (long)input.Length);
            Console.WriteLine(workload + ": " + inputs.Length + " input(s), " + bytes.ToString("N0", System.Globalization.CultureInfo.InvariantCulture) + " distinct characters");
            foreach (IYamlReferenceLibrary library in new[] { YamlReferenceLibraries.YamlDotNet, YamlReferenceLibraries.SharpYaml }) {
                failures += Check(library.Name + " parse", () => YamlWorkloadValidation.ValidateParse(library, workload, inputs));
                failures += Check(library.Name + " load", () => YamlWorkloadValidation.ValidateLoad(library, workload, inputs));
                failures += Check(library.Name + " load-and-save", () => YamlWorkloadValidation.ValidateLoadAndSave(library, workload, inputs));
                failures += Check(library.Name + " emit", () => YamlWorkloadValidation.ValidateEmit(library, workload, inputs, inputs.Select(library.CaptureEvents).ToArray()));
            }
        }

        return failures == 0 ? 0 : 1;
    }

    private static int Check(string name, Action validate) {
        try {
            validate();
            Console.WriteLine("  ok    " + name);
            return 0;
        } catch (Exception exception) {
            Console.WriteLine("  FAIL  " + name + ": " + exception.Message);
            return 1;
        }
    }
}
