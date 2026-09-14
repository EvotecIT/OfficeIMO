namespace OfficeIMO.Html.Runtime;

/// <summary>Caller-supplied planning callback over provider-neutral observations and tools.</summary>
public delegate Task<HtmlAutomationPlannerDecision> HtmlAutomationPlanner(HtmlAutomationTurn turn, CancellationToken cancellationToken);

/// <summary>One bounded planner input.</summary>
public sealed class HtmlAutomationTurn {
    /// <summary>Zero-based planning step.</summary>
    public int Step { get; init; }
    /// <summary>Current page observation.</summary>
    public HtmlPageObservation Observation { get; init; } = null!;
    /// <summary>Results from the preceding decision.</summary>
    public IReadOnlyList<HtmlAutomationToolResult> PreviousResults { get; init; } = Array.Empty<HtmlAutomationToolResult>();
    /// <summary>Available model-SDK-neutral tool definitions.</summary>
    public IReadOnlyList<HtmlAutomationToolDefinition> Tools { get; init; } = Array.Empty<HtmlAutomationToolDefinition>();
}

/// <summary>Planner decision to complete or issue a bounded batch of tool calls.</summary>
public sealed class HtmlAutomationPlannerDecision {
    /// <summary>Whether the workflow has reached its goal.</summary>
    public bool IsComplete { get; init; }
    /// <summary>Optional final planner message.</summary>
    public string? Message { get; init; }
    /// <summary>Tool calls to execute before the next observation.</summary>
    public IReadOnlyList<HtmlAutomationToolCall> Calls { get; init; } = Array.Empty<HtmlAutomationToolCall>();

    /// <summary>Creates a completed decision.</summary>
    public static HtmlAutomationPlannerDecision Complete(string? message = null) => new() { IsComplete = true, Message = message };
    /// <summary>Creates a decision containing one or more tool calls.</summary>
    public static HtmlAutomationPlannerDecision Execute(params HtmlAutomationToolCall[] calls) => new() { Calls = calls };
}

/// <summary>Bounds for a planner-driven observation/action loop.</summary>
public sealed class HtmlAutomationRunOptions {
    /// <summary>Maximum planner decisions.</summary>
    public int MaxSteps { get; set; } = 32;
    /// <summary>Maximum calls accepted in one decision.</summary>
    public int MaxCallsPerStep { get; set; } = 8;
    /// <summary>Observation request used before each decision.</summary>
    public HtmlPageObservationRequest Observation { get; set; } = new() { ActionableOnly = true };

    internal HtmlAutomationRunOptions Snapshot(int maximumOutputCharacters) {
        if (MaxSteps <= 0 || MaxSteps > 10_000) throw new ArgumentOutOfRangeException(nameof(MaxSteps));
        if (MaxCallsPerStep <= 0 || MaxCallsPerStep > 256) throw new ArgumentOutOfRangeException(nameof(MaxCallsPerStep));
        return new HtmlAutomationRunOptions {
            MaxSteps = MaxSteps,
            MaxCallsPerStep = MaxCallsPerStep,
            Observation = (Observation ?? throw new ArgumentNullException(nameof(Observation))).Snapshot(maximumOutputCharacters)
        };
    }
}

/// <summary>Completed planner run with the final observation and operation evidence.</summary>
public sealed class HtmlAutomationRunResult {
    /// <summary>Whether the planner explicitly completed.</summary>
    public bool IsComplete { get; init; }
    /// <summary>Number of planner decisions.</summary>
    public int Steps { get; init; }
    /// <summary>Final planner message.</summary>
    public string? Message { get; init; }
    /// <summary>Latest page observation.</summary>
    public HtmlPageObservation FinalObservation { get; init; } = null!;
    /// <summary>All tool results in execution order.</summary>
    public IReadOnlyList<HtmlAutomationToolResult> ToolResults { get; init; } = Array.Empty<HtmlAutomationToolResult>();
}

/// <summary>Runs a replaceable planner over the same deterministic page tools available to ordinary .NET callers.</summary>
public sealed class HtmlAutomationRunner {
    private readonly HtmlAutomationToolDispatcher _dispatcher;

    /// <summary>Creates a runner with the built-in dispatcher.</summary>
    public HtmlAutomationRunner() : this(new HtmlAutomationToolDispatcher()) { }

    /// <summary>Creates a runner with a caller-selected dispatcher.</summary>
    public HtmlAutomationRunner(HtmlAutomationToolDispatcher dispatcher) =>
        _dispatcher = dispatcher ?? throw new ArgumentNullException(nameof(dispatcher));

    /// <summary>Runs until the planner completes or the configured step bound is reached.</summary>
    public async Task<HtmlAutomationRunResult> RunAsync(IHtmlRuntimePage page, HtmlAutomationPlanner planner,
        HtmlAutomationRunOptions? options = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(page);
        ArgumentNullException.ThrowIfNull(planner);
        HtmlAutomationRunOptions limits = (options ?? new HtmlAutomationRunOptions()).Snapshot(int.MaxValue);
        var allResults = new List<HtmlAutomationToolResult>();
        IReadOnlyList<HtmlAutomationToolResult> previous = Array.Empty<HtmlAutomationToolResult>();
        HtmlPageObservation observation = await page.ObserveAsync(limits.Observation, cancellationToken).ConfigureAwait(false);
        for (int step = 0; step < limits.MaxSteps; step++) {
            HtmlAutomationPlannerDecision decision = await planner(new HtmlAutomationTurn {
                Step = step,
                Observation = observation,
                PreviousResults = previous,
                Tools = HtmlAutomationToolCatalog.GetDefinitions()
            }, cancellationToken).ConfigureAwait(false) ?? throw new InvalidOperationException("The planner returned no decision.");
            if (decision.IsComplete) return new HtmlAutomationRunResult {
                IsComplete = true, Steps = step + 1, Message = decision.Message,
                FinalObservation = observation, ToolResults = Array.AsReadOnly(allResults.ToArray())
            };
            if (decision.Calls == null || decision.Calls.Count == 0 || decision.Calls.Count > limits.MaxCallsPerStep)
                throw new InvalidOperationException("A continuing planner decision must contain a bounded nonempty call list.");
            if (decision.Calls.Select(call => call?.Id).Any(string.IsNullOrWhiteSpace)
                || decision.Calls.Select(call => call.Id).Distinct(StringComparer.Ordinal).Count() != decision.Calls.Count)
                throw new InvalidOperationException("Tool call ids must be nonempty and unique within a decision.");
            var results = new List<HtmlAutomationToolResult>(decision.Calls.Count);
            foreach (HtmlAutomationToolCall call in decision.Calls) {
                HtmlAutomationToolResult result = await _dispatcher.ExecuteAsync(page, call, cancellationToken).ConfigureAwait(false);
                results.Add(result);
                allResults.Add(result);
                if (!result.IsSuccess) break;
            }
            previous = Array.AsReadOnly(results.ToArray());
            observation = previous.LastOrDefault(result => result.Observation != null)?.Observation
                ?? await page.ObserveAsync(limits.Observation, cancellationToken).ConfigureAwait(false);
        }
        return new HtmlAutomationRunResult {
            Steps = limits.MaxSteps,
            Message = "The planner reached its configured step limit.",
            FinalObservation = observation,
            ToolResults = Array.AsReadOnly(allResults.ToArray())
        };
    }
}
