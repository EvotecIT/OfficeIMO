using OfficeIMO.IWork.Benchmarks;

bool phaseCancellation = args.Length == 1 && args[0] == "--validate-phase-cancellation";
bool nativePhaseCancellation = args.Length == 2 && args[0] == "--validate-native-phase-cancellation";
if (!phaseCancellation && !nativePhaseCancellation && (args.Length != 2 || args[0] != "--validate")) {
    Console.Error.WriteLine("Use --validate <IWorkCorpus directory> to validate pinned native operations and active I/O cancellation.");
    Console.Error.WriteLine("Use --validate-phase-cancellation to opt into timer-dependent projection/encoding cancellation checks.");
    Console.Error.WriteLine("Use --validate-native-phase-cancellation <IWorkCorpus directory> for timed Apple-produced source projection cancellation.");
    return 2;
}
try {
    if (phaseCancellation || nativePhaseCancellation) {
        var workload = nativePhaseCancellation ? IWorkRuntimeWorkload.OpenLargeNativeKeynote(args[1])
            : IWorkRuntimeWorkload.CreateNativeKeynote(1000);
        string[] operations = nativePhaseCancellation ? ["CancelDuringProjection"]
            : ["CancelDuringProjection", "CancelDuringNativeEncoding"];
        foreach (string operation in operations) {
            workload.Prepare(operation); workload.Execute(operation); workload.Validate();
            Console.WriteLine($"Keynote/{(nativePhaseCancellation ? "NativeLarge" : "NativeCreated")}/{operation}: cancellation observed; recovered {workload.CancellationRecoveryUnits} slides; destination unchanged");
            workload.ReleaseResults();
        }
        return 0;
    }
    foreach (string kind in new[] { "Pages", "Numbers", "Keynote" }) {
        var workload = new IWorkRuntimeWorkload(kind, args[1]);
        foreach (string operation in new[] { "LoadProject", "ConvertSave", "CancelDuringLoad", "CancelDuringConvert" }) {
            workload.Prepare(operation);
            workload.Execute(operation);
            workload.Validate();
            Console.WriteLine($"{kind}/{operation}: verified {workload.VerifiedUnits}; I/O before cancellation {workload.CancellationProcessedBytes} bytes; accepted padding omissions {workload.KnownPaddingOmissionCount}; known Numbers approximations {workload.KnownNumbersApproximationCount}");
            workload.ReleaseResults();
        }
        if (kind == "Keynote") {
            workload.Prepare("CancelDuringNativeCopy");
            workload.Execute("CancelDuringNativeCopy");
            workload.Validate();
            Console.WriteLine($"{kind}/CancelDuringNativeCopy: verified {workload.VerifiedUnits}; written {workload.CancellationProcessedBytes} bytes");
        }
    }
    var largeApple = IWorkRuntimeWorkload.OpenLargeNativeKeynote(args[1]);
    foreach (string operation in new[] { "LoadProject", "ConvertSave" }) {
        largeApple.Prepare(operation); largeApple.Execute(operation); largeApple.Validate();
        Console.WriteLine($"Keynote/NativeLarge/{operation}: verified {largeApple.VerifiedUnits}; input {largeApple.InputBytes}; output {largeApple.OutputBytes}");
        largeApple.ReleaseResults();
    }
    var largeNative = IWorkRuntimeWorkload.CreateNativeKeynote(1000);
    foreach (string operation in new[] { "LoadProject", "ConvertSave", "NativeSave" }) {
        largeNative.Prepare(operation);
        largeNative.Execute(operation);
        largeNative.Validate();
        Console.WriteLine($"Keynote/NativeCreated/{operation}: verified {largeNative.VerifiedUnits}; input {largeNative.InputBytes}; output {largeNative.OutputBytes}");
        largeNative.ReleaseResults();
    }
    return 0;
} catch (Exception exception) {
    Console.Error.WriteLine(exception);
    return 1;
}
