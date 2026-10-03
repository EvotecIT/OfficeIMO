using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.AI;
using OfficeIMO.Reader;

/// <summary>Request packing with no model inference; PowerForge owns measurement policy and output.</summary>
public sealed class AiPackingWorkload {
    private readonly OfficeAiDocument _document;
    private readonly PackingExecutor _executor = new PackingExecutor();
    private readonly OfficeAiEngine _engine;
    private readonly OfficeAiRequest _request;
    private OfficeAiResult _result;
    public long AllocatedBytes { get; private set; }
    public int MeasurementCalls => _executor.MeasurementCalls;
    public int RequestCount => _result == null ? 0 : _result.RequestCount;

    public AiPackingWorkload(int blocks, int requestCharacters) {
        _document = OfficeAiDocument.FromReadResult(new byte[] { 1 }, new OfficeDocumentReadResult {
            Blocks = Enumerable.Range(0, blocks).Select(index => new OfficeDocumentBlock {
                Id = "block-" + index, Kind = "paragraph", Text = "Evidence fragment number " + index + ".",
                Location = new ReaderLocation { Page = 1 }
            }).ToArray()
        });
        _engine = new OfficeAiEngine(_executor);
        _request = new OfficeAiRequest { Instruction = "Review", Limits = new OfficeAiLimits {
            MaxRequestCharacters = requestCharacters, MaxRequests = 256
        } };
    }

    public void Run() {
        _executor.Requests.Clear(); _executor.MeasurementCalls = 0;
        long allocated = GC.GetTotalAllocatedBytes(true);
        _result = _engine.RunAsync(_document, _request).GetAwaiter().GetResult();
        AllocatedBytes = GC.GetTotalAllocatedBytes(true) - allocated;
    }

    public void Validate() {
        if (_result == null || _result.OmittedEvidenceIds.Count != 0
            || _result.ProcessedEvidenceIds.Count != _document.Evidence.Count
            || _result.RequestCount != _executor.Requests.Count)
            throw new InvalidDataException("Packing did not cover the complete source.");
        var evidence = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (OfficeAiExecutionRequest request in _executor.Requests) {
            if (PackingExecutor.Size(request) > _request.Limits.MaxRequestCharacters)
                throw new InvalidDataException("Packing exceeded the transport request budget.");
            using (JsonDocument input = JsonDocument.Parse(request.InputJson))
                foreach (JsonElement item in input.RootElement.GetProperty("evidence").EnumerateArray())
                    evidence.Add(item.GetProperty("id").GetString(), item.GetProperty("text").GetString());
        }
        if (_document.Evidence.Any(item => !evidence.TryGetValue(item.Id, out string text) || text != item.Text))
            throw new InvalidDataException("Packed evidence differs from the captured source.");
    }

    private sealed class PackingExecutor : IOfficeAiExecutor {
        public List<OfficeAiExecutionRequest> Requests { get; } = new List<OfficeAiExecutionRequest>();
        public int MeasurementCalls;
        public OfficeAiExecutionProfile Profile { get; } = new OfficeAiExecutionProfile {
            Id = "packing", Provider = "no-inference", Model = "no-inference", IsLocal = true, MaxRequestCharacters = 2_000_000
        };
        public static int Size(OfficeAiExecutionRequest request) => checked(request.Instructions.Length + request.InputJson.Length + request.OutputSchema.Length);
        public int MeasureRequestCharacters(OfficeAiExecutionRequest request) { MeasurementCalls++; return Size(request); }
        public Task<OfficeAiExecutionResponse> ExecuteAsync(OfficeAiExecutionRequest request, CancellationToken cancellationToken = default) {
            Requests.Add(request);
            return Task.FromResult(new OfficeAiExecutionResponse("{\"claims\":[],\"fields\":[],\"blocks\":[],\"tables\":[]}"));
        }
    }
}
