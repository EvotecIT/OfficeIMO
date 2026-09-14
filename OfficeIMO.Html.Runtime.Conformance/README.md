# OfficeIMO.Html.Runtime.Conformance

Run the same black-box application workflow against any
`IHtmlRuntimeHost`. The suite checks navigation and history replacement, bounded
observation with password-value omission, revision-bound actions, CSS locators,
keyboard modifiers, delayed and missing-target structured waits, JSON extraction,
stale-reference rejection and inert document capture. It uses no test framework
and does not inspect provider-native browser or interpreter objects.

```csharp
using OfficeIMO.Html.Runtime.Conformance;

HtmlRuntimeConformanceReport report =
    await HtmlRuntimeConformanceSuite.RunAsync(runtimeHost, cancellationToken);

if (!report.Passed) {
    foreach (HtmlRuntimeConformanceCaseResult failure in
             report.Cases.Where(item => !item.Passed)) {
        Console.Error.WriteLine($"{failure.Id}: {failure.Error}");
    }
}
```

Implement another provider through `IHtmlRuntimeHost`, `IHtmlRuntimeContext` and
`IHtmlRuntimePage`. Return owned observation, action, trace and capture models;
keep native pages, element handles, JavaScript values and browser contexts inside
the adapter. Advertise only capabilities that the adapter implements, validate
requests with their public `Snapshot` methods, and reject stale observed
references before resolving or acting on a current provider element.

This suite verifies a controlled common contract. It does not claim complete
browser, JavaScript, accessibility or web-platform compatibility for a provider.
