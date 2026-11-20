using System.Text;

namespace OfficeIMO.Toml.Benchmarks;

/// <summary>Checks that every library sees the same data in every workload and that the savers reproduce their input.</summary>
public static class TomlWorkloadCheck {
    public static int Run() {
        int failures = 0;
        foreach (TomlWorkload workload in Enum.GetValues<TomlWorkload>()) {
            try {
                Validate(workload);
                RecordParse(workload);
                RecordLoad(workload);
                RecordLoadAndSave(workload);
                Console.WriteLine(workload + ": ok");
            } catch (Exception ex) {
                failures++;
                Console.WriteLine(workload + ": " + ex.Message);
            }
        }
        foreach (TomlSerializerWorkload workload in Enum.GetValues<TomlSerializerWorkload>()) {
            try {
                TomlSerializerLibraries.Validate(workload);
                Console.WriteLine("Serializer " + workload + ": ok");
            } catch (Exception ex) {
                failures++;
                Console.WriteLine("Serializer " + workload + ": " + ex.Message);
            }
        }
        return failures == 0 ? 0 : 1;
    }

    public static void Validate(TomlWorkload workload) {
        foreach (string input in TomlWorkloads.Get(workload)) {
            long parsed = TomlLibraries.ParseOfficeIMO(input);
            Expect(parsed, TomlLibraries.ParseTomlyn(input), "Tomlyn parse values");
            long loaded = TomlLibraries.CountOfficeIMO(TomlLibraries.LoadOfficeIMO(input).Root);
            Expect(parsed, loaded, "OfficeIMO load values");
            Expect(loaded, TomlLibraries.CountTomlyn(TomlLibraries.LoadTomlyn(input)), "Tomlyn load values");
            Expect(loaded, TomlLibraries.CountTomlet(TomlLibraries.LoadTomlet(input)), "Tomlet load values");
            Expect(loaded, TomlLibraries.CountCsToml(TomlLibraries.LoadCsToml(Encoding.UTF8.GetBytes(input))), "CsToml load values");
            if (TomlLibraries.SaveOfficeIMO(input) != input) throw new InvalidOperationException("OfficeIMO did not reproduce the input.");
            if (TomlLibraries.SaveTomlyn(input) != input) throw new InvalidOperationException("Tomlyn did not reproduce the input.");
        }
    }

    /// <summary>
    /// The parse lanes report the same keys and decoded values in the same order, and those values are the ones Tomlet's
    /// document model (a third, independent reader) holds.
    /// </summary>
    public static void RecordParse(TomlWorkload workload) {
        string[] inputs = TomlWorkloads.Get(workload);
        List<string>[] officeIMO = inputs.Select(TomlContentCheck.EventsOfficeIMO).ToArray();
        List<string>[] tomlyn = inputs.Select(TomlContentCheck.EventsTomlyn).ToArray();
        List<string>[] expected = inputs.Select(input => TomlContentCheck.ScalarsOf(TomlContentCheck.LeavesTomlet(TomlLibraries.LoadTomlet(input)))).ToArray();
        string detail = "Workload=" + workload + "; " + officeIMO.Sum(events => TomlContentCheck.KeyedValues(events).Count) + " values with their keys in document order, identical across OfficeIMO and Tomlyn (hash " +
            TomlContentCheck.Hash(officeIMO.SelectMany(TomlContentCheck.KeyedValues)) + "); " + expected.Sum(values => values.Count) + " decoded values match Tomlet's document";
        TomlContentCheck.Record("OfficeIMO", "Default", detail, () => CompareEach(inputs, index =>
            TomlContentCheck.Compare("values against Tomlet", expected[index], TomlContentCheck.ScalarsOfEvents(officeIMO[index])) ??
            TomlContentCheck.Compare("keyed values against Tomlyn", TomlContentCheck.KeyedValues(tomlyn[index]), TomlContentCheck.KeyedValues(officeIMO[index]))));
        TomlContentCheck.Record("Tomlyn", "Default", detail, () => CompareEach(inputs, index =>
            TomlContentCheck.Compare("values against Tomlet", expected[index], TomlContentCheck.ScalarsOfEvents(tomlyn[index])) ??
            TomlContentCheck.Compare("keyed values against OfficeIMO", TomlContentCheck.KeyedValues(officeIMO[index]), TomlContentCheck.KeyedValues(tomlyn[index]))));
    }

    /// <summary>
    /// Every loader's model holds the same leaf values at the same paths. Tomlyn's model is the reference for the other
    /// three and CsToml's for Tomlyn. OfficeIMO's LazyValues lane is not recorded: it leaves values undecoded, so only
    /// its Default lane, which decodes every value, does the reference loaders' work.
    /// </summary>
    public static void RecordLoad(TomlWorkload workload) {
        string[] inputs = TomlWorkloads.Get(workload);
        List<string>[] tomlet = inputs.Select(input => TomlContentCheck.LeavesTomlet(TomlLibraries.LoadTomlet(input))).ToArray();
        List<string>[] tomlyn = inputs.Select(input => TomlContentCheck.LeavesTomlyn(TomlLibraries.LoadTomlyn(input))).ToArray();
        List<string>[] csToml = inputs.Select(input => TomlContentCheck.LeavesCsToml(TomlLibraries.LoadCsToml(Encoding.UTF8.GetBytes(input)))).ToArray();
        List<string>[] officeIMO = inputs.Select(input => TomlContentCheck.LeavesOfficeIMO(TomlLibraries.LoadOfficeIMO(input))).ToArray();
        string detail = "Workload=" + workload + "; " + tomlyn.Sum(leaves => leaves.Count) + " leaf values with their table and array paths (hash " +
            TomlContentCheck.Hash(tomlyn.SelectMany(leaves => leaves)) + ")";
        TomlContentCheck.Record("OfficeIMO", "Default", detail + " match Tomlyn's model", () => CompareEach(inputs, index => TomlContentCheck.Compare("leaves against Tomlyn", tomlyn[index], officeIMO[index])));
        TomlContentCheck.Record("Tomlet", "Default", detail + " match Tomlyn's model", () => CompareEach(inputs, index => TomlContentCheck.Compare("leaves against Tomlyn", tomlyn[index], tomlet[index])));
        TomlContentCheck.Record("CsToml", "Default", detail + " match Tomlyn's model", () => CompareEach(inputs, index => TomlContentCheck.Compare("leaves against Tomlyn", tomlyn[index], csToml[index])));
        TomlContentCheck.Record("Tomlyn", "Default", detail + " match CsToml's model", () => CompareEach(inputs, index => TomlContentCheck.Compare("leaves against CsToml", csToml[index], tomlyn[index])));
    }

    /// <summary>Both savers return the input text unchanged.</summary>
    public static void RecordLoadAndSave(TomlWorkload workload) {
        string[] inputs = TomlWorkloads.Get(workload);
        string detail = "Workload=" + workload + "; saved text identical to the input (" + inputs.Sum(input => (long)input.Length) + " characters in " + inputs.Length + " files)";
        TomlContentCheck.Record("OfficeIMO", "Default", detail, () => inputs.All(input => TomlLibraries.SaveOfficeIMO(input) == input) ? null : "the saved text differs from the input");
        TomlContentCheck.Record("Tomlyn", "Default", detail, () => inputs.All(input => TomlLibraries.SaveTomlyn(input) == input) ? null : "the saved text differs from the input");
    }

    private static string? CompareEach(string[] inputs, Func<int, string?> compare) {
        var failures = new List<string>();
        for (int index = 0; index < inputs.Length; index++) {
            string? difference = compare(index);
            if (difference is not null) failures.Add("input " + (index + 1) + " of " + inputs.Length + ": " + difference);
        }
        return failures.Count == 0 ? null : string.Join(" | ", failures);
    }

    private static void Expect(long expected, long actual, string what) {
        if (expected != actual) throw new InvalidOperationException(what + ": expected " + expected + " but got " + actual + ".");
    }
}
