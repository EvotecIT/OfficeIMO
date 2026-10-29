using System.Text.Json;

namespace OfficeIMO.Publisher.WriterProbe;

internal static class Program {
    internal static readonly JsonSerializerOptions JsonOptions = new() {
        WriteIndented = true,
        Converters = { new System.Text.Json.Serialization.JsonStringEnumConverter() }
    };

    public static int Main(string[] args) {
        try {
            if (args.Length is >= 2 and <= 3 && args[0] == "inventory") {
                string json = JsonSerializer.Serialize(NativeInventory.Read(args[1]), JsonOptions);
                if (args.Length == 3) WriteNew(args[2], System.Text.Encoding.UTF8.GetBytes(json));
                else Console.WriteLine(json);
                return 0;
            }
            if (args.Length is >= 3 and <= 4 && args[0] == "create") {
                byte[] bytes = CandidateWriter.Create(args[1], args.Length == 4 ? args[3] : null);
                WriteNew(args[2], bytes);
                Console.WriteLine(JsonSerializer.Serialize(new {
                    schema = "officeimo.publisher.writer-probe.candidate.v1",
                    profile = "unqualified-native-schema-quill-default-font-v9",
                    hypothesis = args[0],
                    kind = args[1],
                    rootClassId = CandidateWriter.PublicationClassId,
                    pageGraph = new {
                        orderedReferences = NativePublicationGraph.OrderedPageCount,
                        dpgMaster = NativePublicationGraph.MasterPageCount,
                        dpgSpecial = NativePublicationGraph.SpecialPageBoundary,
                        firstPhysicalIndex = NativePublicationGraph.MasterPageCount,
                        lastPhysicalIndex = NativePublicationGraph.OrderedPageCount - NativePublicationGraph.SpecialPageBoundary
                    },
                    output = Path.GetFullPath(args[2]),
                    bytes = bytes.Length,
                    sha256 = NativeInventory.Hash(bytes),
                    sourcePublication = (string?)null,
                    nativeAcceptance = "not-run",
                    note = "Template-free synthesis hypothesis; acceptance by OfficeIMO does not establish acceptance by Microsoft Publisher."
                }, JsonOptions));
                return 0;
            }
            if (args.Length == 3 && args[0] == "render") {
                Console.WriteLine(JsonSerializer.Serialize(ManagedComparison.Render(args[1], args[2]), JsonOptions));
                return 0;
            }
            if (args.Length is 3 or 4 && args[0] == "diff") {
                string json = JsonSerializer.Serialize(NativeDiff.Compare(args[1], args[2]), JsonOptions);
                if (args.Length == 4) WriteNew(args[3], System.Text.Encoding.UTF8.GetBytes(json));
                else Console.WriteLine(json);
                return 0;
            }
            Console.Error.WriteLine("inventory <input.pub> [report.json]\ndiff <left.pub> <right.pub> [report.json]\ncreate <blank|text|shape|image> <candidate.pub> [text-or-image-path]\nrender <input.pub> <unused-output-directory>");
            return 2;
        } catch (Exception exception) {
            Console.Error.WriteLine(exception.GetType().Name + ": " + exception.Message);
            return 1;
        }
    }

    private static void WriteNew(string path, byte[] bytes) {
        using var stream = new FileStream(Path.GetFullPath(path), FileMode.CreateNew, FileAccess.Write);
        stream.Write(bytes);
    }
}
