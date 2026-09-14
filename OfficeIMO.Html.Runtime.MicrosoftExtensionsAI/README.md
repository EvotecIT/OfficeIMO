# OfficeIMO.Html.Runtime.MicrosoftExtensionsAI

Optional `Microsoft.Extensions.AI` integration for the provider-neutral
`HtmlAutomationRunner`. Applications supply their own `IChatClient`, messages,
model settings, credentials, approval policy and retry behavior. The adapter
only translates the published OfficeIMO tool declarations and returned function
calls into bounded planner decisions.

```csharp
using Microsoft.Extensions.AI;
using OfficeIMO.Html.Runtime;
using OfficeIMO.Html.Runtime.MicrosoftExtensionsAI;

IChatClient client = CreateApplicationChatClient();
var planner = new HtmlAutomationChatPlanner(client, turn => new HtmlAutomationChatRequest {
    Messages = new[] {
        new ChatMessage(ChatRole.System,
            "Use the supplied tools to complete the bounded page workflow."),
        new ChatMessage(ChatRole.User,
            $"Goal: approve the report. Current page observation: {HtmlRuntimeJson.Serialize(turn.Observation)}")
    },
    Options = new ChatOptions { ModelId = "application-selected-model" }
});

HtmlAutomationRunResult result = await new HtmlAutomationRunner().RunAsync(
    page,
    planner.CreatePlanner(),
    new HtmlAutomationRunOptions { MaxSteps = 8, MaxCallsPerStep = 2 },
    cancellationToken);
```

The planner adds declaration-only `observe`, `act`, `navigate` and `capture`
tools using the exact OfficeIMO JSON Schemas and makes one `IChatClient` call per
runner turn. Returned function calls go back through `HtmlAutomationRunner`, so
the existing step, call, observation and runtime limits remain authoritative.
Text-only output completes the run. The adapter does not dispose the client,
retry a model call, execute provider-native handles or expand the trusted-content
boundary. `HtmlAutomationChatPlannerOptions` also bounds calls, UTF-8 argument
bytes, nested depth and item count before a planner decision is returned. A
caller-supplied function declaration using an OfficeIMO-reserved tool name is
rejected so the published schema cannot be replaced accidentally. Function calls
that the model provider could not map are rejected before their argument object
can become an OfficeIMO tool call.
