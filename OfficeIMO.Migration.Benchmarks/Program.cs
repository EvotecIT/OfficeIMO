using OfficeIMO.Migration.Benchmarks;

if (args.Length != 3 || args[0] != "--validate" || !int.TryParse(args[2], out int units)) {
    Console.Error.WriteLine("Use --validate <Word69|Biff4|IWork09|Dbx|OfficeXml> <units>. PowerForge owns timing.");
    return 2;
}
try {
    MigrationWorkload workload = MigrationWorkload.Create(args[1], units);
    string[] operations = args[1] == "Dbx" ? ["LoadInspect", "ConvertSave", "SearchContent"] : ["LoadInspect", "ConvertSave"];
    foreach (string operation in operations) {
        try {
            workload.Execute(operation);
            workload.Validate();
            if (workload.VerifiedUnits != units) throw new InvalidDataException("Incomplete workload validation.");
            Console.WriteLine(System.Text.Json.JsonSerializer.Serialize(new {
                Profile = args[1], Operation = operation, Units = units, workload.InputBytes,
                workload.InputSha256, workload.OutputBytes, workload.VerifiedUnits, workload.DiagnosticDigest
            }));
        } finally { workload.ReleaseResults(); }
    }
    return 0;
} catch (Exception exception) {
    Console.Error.WriteLine(exception);
    return 1;
}

