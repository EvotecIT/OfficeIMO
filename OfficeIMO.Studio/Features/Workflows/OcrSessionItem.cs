using CommunityToolkit.Mvvm.ComponentModel;
using OfficeIMO.Studio.Infrastructure.Localization;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Workflows;

/// <summary>One selected source and its most recent OCR session attempt.</summary>
public sealed partial class OcrSessionItem : ObservableObject {
    private readonly IStudioLocalizer _localizer;
    internal OcrSessionItem(string source, string name, IStudioLocalizer localizer, IEnumerable<string>? reservedOutputNames = null) {
        InputPath = source; Name = name; _localizer = localizer;
        IsPdf = string.Equals(Path.GetExtension(name), ".pdf", StringComparison.OrdinalIgnoreCase);
        OutputName = IsPdf ? Path.GetFileNameWithoutExtension(name) + "-searchable.pdf" : name + ".txt";
        var reserved = new HashSet<string>(reservedOutputNames ?? [], StringComparer.OrdinalIgnoreCase);
        string stem = Path.GetFileNameWithoutExtension(OutputName);
        string extension = Path.GetExtension(OutputName);
        for (int suffix = 2; reserved.Contains(OutputName); suffix++) OutputName = $"{stem} ({suffix}){extension}";
    }
    public string Id { get; } = Guid.NewGuid().ToString("N");
    public string InputPath { get; }
    public string Name { get; }
    public bool IsPdf { get; }
    public string OutputName { get; }
    public string Route => IsPdf ? _localizer.GetOrDefault("OcrSession.PdfRoute", "PDF → searchable PDF")
        : _localizer.GetOrDefault("OcrSession.ImageRoute", "Image → reviewed text");
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(StatusLabel))]
    [NotifyPropertyChangedFor(nameof(CanRetry))]
    [NotifyPropertyChangedFor(nameof(HasOutput))]
    private OfficeWorkflowStatus? _status;
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(StatusLabel))]
    private bool _isRunning;
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasOutput))]
    private string? _outputPath;
    [ObservableProperty] private string? _summary;
    [ObservableProperty] private OfficeWorkflowOutputRecovery? _recovery;
    public bool CanRetry => Status is OfficeWorkflowStatus.Failed or OfficeWorkflowStatus.Cancelled;
    public bool HasOutput => Status == OfficeWorkflowStatus.Completed && OutputPath is not null;
    public string StatusLabel => IsRunning ? _localizer.GetOrDefault("Jobs.Running", "Running")
        : Status is null ? _localizer.GetOrDefault("Jobs.Queued", "Queued")
        : Status == OfficeWorkflowStatus.Unconfirmed ? _localizer.GetOrDefault("Jobs.Unconfirmed", "Check output")
        : _localizer.GetOrDefault("Workflow.Status." + Status, Status.ToString()!);
    internal void ResetAttempt() { Status = null; IsRunning = false; OutputPath = null; Summary = null; Recovery = null; }
    internal void Apply(OfficeOcrSessionResult result) {
        IsRunning = false; Status = result.Status; OutputPath = result.OutputPath; Summary = result.Summary; Recovery = result.Recovery;
    }
}
