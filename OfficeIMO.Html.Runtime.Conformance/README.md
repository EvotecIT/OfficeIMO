# OfficeIMO.Html.Runtime.Conformance

Run the same seven-case, 45-assertion black-box application suite against any
`IHtmlRuntimeHost`. It covers supplied and blocked resources, redirects, console
and policy evidence, navigation and history replacement, bounded observation with
password-value omission, revision-bound actions, CSS locators, keyboard modifiers,
delayed and missing-target structured waits, JSON extraction, stale-reference
recovery, deterministic capture manifests and a complete rules-driven tool
workflow. It uses no test framework and does not inspect provider-native browser
or interpreter objects.

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

HtmlRuntimeQualificationManifest profile =
    HtmlRuntimeQualificationCatalog.Get("web-application-v1");
HtmlRuntimeQualificationResult qualification =
    HtmlRuntimeQualificationCatalog.EvaluateProvider(profile, report);

if (!qualification.Passed) {
    throw new InvalidOperationException("The provider does not match the profile manifest.");
}
```

When a provider implements only one profile, pass its manifest directly to
`HtmlRuntimeConformanceSuite.RunAsync(runtimeHost, profile, cancellationToken)`.
The suite then runs exactly that manifest's cases before evaluation.

The embedded schema-1 catalog defines `scripted-document-v1` (2 cases and 11
assertions), `web-application-v1` (7 cases and 45 assertions), and
`programmatic-automation-v1` (3 cases and 21 assertions). Each profile pins its
selected specification revisions, provider expectations, exclusions, untested
areas, and explicit Test262/WPT counts. No Test262 or WPT subset is selected in
v1; the manifest records that boundary instead of implying upstream-suite
conformance. `EvaluateProvider` reports provider evidence only. For a profile
that declares consumers, pass measured outcomes through `EvaluateConsumer` and
then call `Evaluate` with the provider report and every consumer result. The
combined profile cannot pass when declared consumer evidence is absent.
`HtmlRuntimeConformanceJson.Serialize` exports reports, manifests and provider,
consumer or combined qualification results through source-generated JSON metadata.

Implement another provider through `IHtmlRuntimeHost`, `IHtmlRuntimeContext` and
`IHtmlRuntimePage`. Return owned observation, action, trace and capture models;
keep native pages, element handles, JavaScript values and browser contexts inside
the adapter. Advertise only capabilities that the adapter implements, validate
requests with their public `Snapshot` methods, and reject stale observed
references before resolving or acting on a current provider element.

This suite verifies a controlled common contract. It does not claim complete
browser, JavaScript, accessibility or web-platform compatibility for a provider.
