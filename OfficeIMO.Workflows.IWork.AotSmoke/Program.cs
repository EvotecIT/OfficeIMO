using System.Runtime.CompilerServices;

if (RuntimeFeature.IsDynamicCodeSupported || RuntimeFeature.IsDynamicCodeCompiled)
    throw new InvalidOperationException("The composed iWork PDF qualification must execute as NativeAOT.");
if (args.Length > 1) throw new ArgumentException("Supply at most one empty output directory.");
bool retain = args.Length == 1;
string root = retain ? Path.GetFullPath(args[0])
    : Path.Combine(Path.GetTempPath(), "officeimo-iwork-pdf-aot-" + Guid.NewGuid().ToString("N"));
Directory.CreateDirectory(root);
if (Directory.EnumerateFileSystemEntries(root).Any())
    throw new ArgumentException("The qualification output directory must be empty.");
try {
    await NativePdfScenarios.RunAsync(root);
    Console.WriteLine("PASS | bounded composed iWork PDF NativeAOT | "
        + System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription);
} finally {
    if (!retain) Directory.Delete(root, recursive: true);
}
