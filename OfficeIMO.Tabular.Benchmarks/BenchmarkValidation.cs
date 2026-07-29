namespace OfficeIMO.Tabular.Benchmarks;

internal static class BenchmarkValidation {
    internal static void Run() {
        FixtureData.EnsureAuthentic();
        var scenarios = new[] {
            Pair("CSV string", TabularBenchmarkOperations.ReadSylvanCsvStrings(), TabularBenchmarkOperations.ReadOfficeCsvStrings()),
            Pair("CSV typed", TabularBenchmarkOperations.ReadSylvanCsvTyped(), TabularBenchmarkOperations.ReadOfficeCsvTyped()),
            Pair("XLSX typed", TabularBenchmarkOperations.ReadSylvanXlsxTyped(), TabularBenchmarkOperations.ReadOfficeXlsxTyped()),
            Pair("XLSX binder", TabularBenchmarkOperations.ReadSylvanXlsxRecords(), TabularBenchmarkOperations.ReadOfficeXlsxRecords()),
            Pair("XLSB typed", TabularBenchmarkOperations.ReadSylvanXlsbTyped(), TabularBenchmarkOperations.ReadOfficeXlsbTyped())
        };

        foreach (var scenario in scenarios) {
            ValidateExpectedShape(scenario.Name, scenario.Sylvan);
            ValidateExpectedShape(scenario.Name, scenario.OfficeIMO);
            if (scenario.Sylvan != scenario.OfficeIMO) {
                throw new InvalidDataException(
                    $"{scenario.Name} produced different results. Sylvan={scenario.Sylvan}; OfficeIMO={scenario.OfficeIMO}.");
            }

            Console.WriteLine($"{scenario.Name}: {scenario.OfficeIMO}");
        }
    }

    private static (string Name, Observation Sylvan, Observation OfficeIMO) Pair(
        string name,
        Observation sylvan,
        Observation officeImo) => (name, sylvan, officeImo);

    private static void ValidateExpectedShape(string scenario, Observation observation) {
        int expectedCells = FixtureData.ExpectedRows * FixtureData.ExpectedColumns;
        if (observation.Rows != FixtureData.ExpectedRows || observation.Cells != expectedCells) {
            throw new InvalidDataException(
                $"{scenario} expected {FixtureData.ExpectedRows} rows/{expectedCells} cells, got {observation.Rows} rows/{observation.Cells} cells.");
        }
    }
}
