namespace OfficeIMO.Email;

/// <summary>Typed helpers over the lossless iCalendar component/property model.</summary>
public static class IcsComponentExtensions {
    /// <summary>Gets the first temporal property with the supplied name.</summary>
    public static IcsTemporalValue? GetTemporalValue(this ContentLineComponent component, string propertyName) {
        ArgumentNullException.ThrowIfNull(component);
        return IcsTemporalValue.TryParse(component.GetFirstProperty(propertyName), out IcsTemporalValue value)
            ? value
            : (IcsTemporalValue?)null;
    }

    /// <summary>Creates or replaces one temporal property without resolving TZID through the host OS.</summary>
    public static ContentLineProperty SetTemporalValue(this ContentLineComponent component, string propertyName,
        IcsTemporalValue value) {
        ArgumentNullException.ThrowIfNull(component);
        ContentLineProperty property = component.SetProperty(propertyName, string.Empty);
        value.ApplyTo(property);
        return property;
    }

    /// <summary>Parses every direct RRULE property.</summary>
    public static IEnumerable<IcsRecurrenceRule> GetRecurrenceRules(this ContentLineComponent component) {
        ArgumentNullException.ThrowIfNull(component);
        foreach (ContentLineProperty property in component.GetProperties("RRULE"))
            yield return IcsRecurrenceRule.Parse(property.Value);
    }

    /// <summary>Adds an RRULE property.</summary>
    public static ContentLineProperty AddRecurrenceRule(this ContentLineComponent component,
        IcsRecurrenceRule rule) {
        ArgumentNullException.ThrowIfNull(component);
        ArgumentNullException.ThrowIfNull(rule);
        return component.AddProperty("RRULE", rule.ToString());
    }
}
