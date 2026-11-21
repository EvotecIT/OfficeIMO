namespace OfficeIMO.Yaml.Tests.Conformance;

public sealed partial class YamlConformanceChecks {
    /// <summary>
    /// Compares the events of a valid case with the content of <c>test.event</c>: kinds, scalar values, tags, anchors and
    /// alias names. Unlike <see cref="CheckEvents"/> it applies to engines that do not expose presentation details.
    /// </summary>
    public YamlCheckResult CheckEventContent(YamlTestSuiteCase testCase) {
        if (testCase == null) throw new ArgumentNullException(nameof(testCase));
        const YamlConformanceCheck check = YamlConformanceCheck.EventContent;
        if (testCase.IsError || !Has(YamlEngineFeatures.Events)) return Skip(testCase.Id, check);
        if (!TryRun(() => _engine.Parse(testCase.Input), out IReadOnlyList<YamlTestEvent>? events, out string? error)) {
            return new YamlCheckResult(testCase.Id, check, YamlCheckOutcome.Error, error);
        }

        IReadOnlyList<YamlTestEvent> expected = YamlTestEventParser.Parse(testCase.ExpectedEvents);
        int index = YamlTestEventText.FindContentDifference(expected, events!);
        if (index < 0) return new YamlCheckResult(testCase.Id, check, YamlCheckOutcome.Pass);
        string wanted = index < expected.Count ? expected[index].ToString() : "<end>";
        string actual = index < events!.Count ? events[index].ToString() : "<end>";
        return new YamlCheckResult(testCase.Id, check, YamlCheckOutcome.Fail, "event " + index + ": expected '" + wanted + "' but got '" + actual + "'");
    }
}
