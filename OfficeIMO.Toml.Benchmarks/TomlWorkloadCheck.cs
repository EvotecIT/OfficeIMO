namespace OfficeIMO.Toml.Benchmarks;

/// <summary>Checks that every library sees the same data in every workload, and which savers reproduce their input.</summary>
public static class TomlWorkloadCheck {
    public static int Run() {
        int failures = 0;
        foreach (TomlWorkload workload in Enum.GetValues<TomlWorkload>()) {
            try {
                Validate(workload);
                string? tomlet = TomletSaveLimitation(workload);
                Console.WriteLine(workload + ": ok" + (tomlet is null ? string.Empty : " (Tomlet cannot save it: " + tomlet + ")"));
            } catch (Exception ex) {
                failures++;
                Console.WriteLine(workload + ": " + ex.Message);
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
            if (TomlLibraries.SaveOfficeIMO(input) != input) throw new InvalidOperationException("OfficeIMO did not reproduce the input.");
            if (TomlLibraries.SaveTomlyn(input) != input) throw new InvalidOperationException("Tomlyn did not reproduce the input.");
        }
    }

    /// <summary>Returns why Tomlet cannot re-serialize a workload, or null; its saver rejects some valid keys.</summary>
    public static string? TomletSaveLimitation(TomlWorkload workload) {
        foreach (string input in TomlWorkloads.Get(workload)) {
            try {
                _ = TomlLibraries.SaveTomlet(input);
            } catch (Exception ex) {
                return ex.Message;
            }
        }
        return null;
    }

    private static void Expect(long expected, long actual, string what) {
        if (expected != actual) throw new InvalidOperationException(what + ": expected " + expected + " but got " + actual + ".");
    }
}
