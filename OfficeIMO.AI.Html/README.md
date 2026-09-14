# OfficeIMO.AI.Html

Use an `OfficeIMO.AI.IOfficeAiExecutor` in a .NET 8 or .NET 10 application as the
optional planner for bounded OfficeIMO HTML runtime automation. The HTML runtime
remains deterministic and model-independent; this package translates its owned
observations and tool schemas into one structured OfficeIMO AI request per planning
turn.

Use `OfficeIMO.AI.IntelligenceX` for ChatGPT, Copilot, a local model, or an
OpenAI-compatible endpoint. Applications own the instructions, connection,
credentials, model selection, approval policy, and retry behavior.

```csharp
using OfficeIMO.AI.Html;
using OfficeIMO.Html.Runtime;

var planner = new HtmlAutomationAiPlanner(executor, turn =>
    $"Complete the approved page workflow. Current step: {turn.Step}.");

HtmlAutomationRunResult result = await new HtmlAutomationRunner().RunAsync(
    page,
    planner.CreatePlanner(),
    new HtmlAutomationRunOptions { MaxSteps = 8, MaxCallsPerStep = 2 },
    cancellationToken);
```

The bridge sends only the bounded observation, previous tool results, and exact
OfficeIMO tool declarations. Provider output is treated as untrusted JSON. The
OfficeIMO AI planner rejects truncated responses, undeclared tools, duplicate
call identifiers, oversized arguments, excessive nesting, and malformed
completion decisions. It validates nested arguments against the matching OfficeIMO
schema and removes provider-required nulls for optional values. The HTML runtime then
performs typed validation again before it invokes page behavior.
