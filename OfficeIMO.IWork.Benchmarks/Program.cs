using OfficeIMO.IWork.Benchmarks;

if (args.Length != 2 || args[0] != "--validate") {
    Console.Error.WriteLine("Use --validate <IWorkCorpus directory> to validate pinned native operations and active I/O cancellation.");
    return 2;
}
try {
    foreach (string kind in new[] { "Pages", "Numbers", "Keynote" }) {
        var workload = new IWorkRuntimeWorkload(kind, args[1]);
        foreach (string operation in new[] { "LoadProject", "ConvertSave", "CancelDuringLoad", "CancelDuringConvert" }) {
            workload.Prepare(operation);
            workload.Execute(operation);
            workload.Validate();
            Console.WriteLine($"{kind}/{operation}: verified {workload.VerifiedUnits}; I/O before cancellation {workload.CancellationProcessedBytes} bytes");
            workload.ReleaseResults();
        }
        if (kind == "Keynote") {
            workload.Prepare("CancelDuringNativeCopy");
            workload.Execute("CancelDuringNativeCopy");
            workload.Validate();
            Console.WriteLine($"{kind}/CancelDuringNativeCopy: verified {workload.VerifiedUnits}; written {workload.CancellationProcessedBytes} bytes");
        }
    }
    return 0;
} catch (Exception exception) {
    Console.Error.WriteLine(exception);
    return 1;
}
