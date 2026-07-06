namespace OfficeIMO.Pdf;

public sealed partial class PdfDocumentConversionResult {
    private static void AddOptionalContentIssues(List<PdfConversionProofIssue> issues, PdfDocumentInfo documentInfo, PdfConversionProofOptions options) {
        if (!HasRequiredOptionalContent(options)) {
            return;
        }

        PdfOptionalContentProperties? optionalContent = documentInfo.OptionalContent;
        if (optionalContent is null) {
            AddMissingOptionalContentIssues(issues, options);
            return;
        }

        AddStringIssue(issues, "OptionalContent.DefaultConfigurationName", options.RequiredOptionalContentDefaultConfigurationName, optionalContent.DefaultConfigurationName);
        AddStringIssue(issues, "OptionalContent.DefaultConfigurationCreator", options.RequiredOptionalContentDefaultConfigurationCreator, optionalContent.DefaultConfigurationCreator);
        AddStringIssue(issues, "OptionalContent.BaseState", options.RequiredOptionalContentBaseState, optionalContent.BaseState);

        if (options.RequiredOptionalContentGroupCountAtLeast.HasValue &&
            optionalContent.GroupCount < options.RequiredOptionalContentGroupCountAtLeast.Value) {
            issues.Add(new PdfConversionProofIssue(
                "OptionalContent.GroupCount",
                "at least " + options.RequiredOptionalContentGroupCountAtLeast.Value.ToString(System.Globalization.CultureInfo.InvariantCulture),
                optionalContent.GroupCount.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        }

        AddOptionalContentGroupNameIssues(issues, optionalContent, options.RequiredOptionalContentGroupNames, "OptionalContent.GroupName", GroupNameMatches);
        AddOptionalContentGroupNameIssues(issues, optionalContent, options.RequiredOptionalContentVisibleGroupNames, "OptionalContent.VisibleGroupName", GroupIsInitiallyVisible);
        AddOptionalContentGroupNameIssues(issues, optionalContent, options.RequiredOptionalContentHiddenGroupNames, "OptionalContent.HiddenGroupName", GroupIsInitiallyHidden);
        AddOptionalContentGroupNameIssues(issues, optionalContent, options.RequiredOptionalContentLockedGroupNames, "OptionalContent.LockedGroupName", GroupIsLocked);
        AddOptionalContentGroupNameIssues(issues, optionalContent, options.RequiredOptionalContentOrderedGroupNames, "OptionalContent.OrderedGroupName", GroupIsInDefaultOrder);
    }

    private static void AddMissingOptionalContentIssues(List<PdfConversionProofIssue> issues, PdfConversionProofOptions options) {
        AddStringIssue(issues, "OptionalContent.DefaultConfigurationName", options.RequiredOptionalContentDefaultConfigurationName, null);
        AddStringIssue(issues, "OptionalContent.DefaultConfigurationCreator", options.RequiredOptionalContentDefaultConfigurationCreator, null);
        AddStringIssue(issues, "OptionalContent.BaseState", options.RequiredOptionalContentBaseState, null);

        if (options.RequiredOptionalContentGroupCountAtLeast.HasValue) {
            issues.Add(new PdfConversionProofIssue(
                "OptionalContent.GroupCount",
                "at least " + options.RequiredOptionalContentGroupCountAtLeast.Value.ToString(System.Globalization.CultureInfo.InvariantCulture),
                "missing"));
        }

        AddMissingOptionalContentGroupNameIssues(issues, options.RequiredOptionalContentGroupNames, "OptionalContent.GroupName");
        AddMissingOptionalContentGroupNameIssues(issues, options.RequiredOptionalContentVisibleGroupNames, "OptionalContent.VisibleGroupName");
        AddMissingOptionalContentGroupNameIssues(issues, options.RequiredOptionalContentHiddenGroupNames, "OptionalContent.HiddenGroupName");
        AddMissingOptionalContentGroupNameIssues(issues, options.RequiredOptionalContentLockedGroupNames, "OptionalContent.LockedGroupName");
        AddMissingOptionalContentGroupNameIssues(issues, options.RequiredOptionalContentOrderedGroupNames, "OptionalContent.OrderedGroupName");
    }

    private static void AddOptionalContentGroupNameIssues(
        List<PdfConversionProofIssue> issues,
        PdfOptionalContentProperties optionalContent,
        IList<string> requiredNames,
        string feature,
        Func<PdfOptionalContentGroup, string, bool> predicate) {
        for (int i = 0; i < requiredNames.Count; i++) {
            string name = requiredNames[i];
            if (!ContainsOptionalContentGroup(optionalContent.Groups, name, predicate)) {
                issues.Add(new PdfConversionProofIssue(feature, name, "missing"));
            }
        }
    }

    private static void AddMissingOptionalContentGroupNameIssues(List<PdfConversionProofIssue> issues, IList<string> requiredNames, string feature) {
        for (int i = 0; i < requiredNames.Count; i++) {
            issues.Add(new PdfConversionProofIssue(feature, requiredNames[i], "missing"));
        }
    }

    private static bool ContainsOptionalContentGroup(
        IReadOnlyList<PdfOptionalContentGroup> groups,
        string name,
        Func<PdfOptionalContentGroup, string, bool> predicate) {
        for (int i = 0; i < groups.Count; i++) {
            if (predicate(groups[i], name)) {
                return true;
            }
        }

        return false;
    }

    private static bool GroupNameMatches(PdfOptionalContentGroup group, string name) {
        return string.Equals(group.Name, name, StringComparison.Ordinal);
    }

    private static bool GroupIsInitiallyVisible(PdfOptionalContentGroup group, string name) {
        return GroupNameMatches(group, name) && group.IsInitiallyVisible == true;
    }

    private static bool GroupIsInitiallyHidden(PdfOptionalContentGroup group, string name) {
        return GroupNameMatches(group, name) && group.IsInitiallyVisible == false;
    }

    private static bool GroupIsLocked(PdfOptionalContentGroup group, string name) {
        return GroupNameMatches(group, name) && group.IsLocked;
    }

    private static bool GroupIsInDefaultOrder(PdfOptionalContentGroup group, string name) {
        return GroupNameMatches(group, name) && group.IsInDefaultOrder;
    }

    private static bool HasRequiredOptionalContent(PdfConversionProofOptions options) {
        return options.RequiredOptionalContentDefaultConfigurationName is not null ||
            options.RequiredOptionalContentDefaultConfigurationCreator is not null ||
            options.RequiredOptionalContentBaseState is not null ||
            options.RequiredOptionalContentGroupCountAtLeast.HasValue ||
            options.RequiredOptionalContentGroupNames.Count > 0 ||
            options.RequiredOptionalContentVisibleGroupNames.Count > 0 ||
            options.RequiredOptionalContentHiddenGroupNames.Count > 0 ||
            options.RequiredOptionalContentLockedGroupNames.Count > 0 ||
            options.RequiredOptionalContentOrderedGroupNames.Count > 0;
    }
}
