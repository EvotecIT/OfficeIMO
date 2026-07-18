namespace OfficeIMO.Email;

/// <summary>Typed helpers over the lossless iCalendar component/property model.</summary>
public static class IcsComponentExtensions {
    /// <summary>Gets the first temporal property with the supplied name.</summary>
    public static IcsTemporalValue? GetTemporalValue(this ContentLineComponent component, string propertyName) {
        if (component == null) throw new ArgumentNullException(nameof(component));
        return IcsTemporalValue.TryParse(component.GetFirstProperty(propertyName), out IcsTemporalValue value)
            ? value
            : (IcsTemporalValue?)null;
    }

    /// <summary>Creates or replaces one temporal property without resolving TZID through the host OS.</summary>
    public static ContentLineProperty SetTemporalValue(this ContentLineComponent component, string propertyName,
        IcsTemporalValue value) {
        if (component == null) throw new ArgumentNullException(nameof(component));
        ContentLineProperty property = component.SetProperty(propertyName, string.Empty);
        value.ApplyTo(property);
        return property;
    }

    /// <summary>Parses every direct RRULE property.</summary>
    public static IEnumerable<IcsRecurrenceRule> GetRecurrenceRules(this ContentLineComponent component) {
        if (component == null) throw new ArgumentNullException(nameof(component));
        foreach (ContentLineProperty property in component.GetProperties("RRULE"))
            yield return IcsRecurrenceRule.Parse(property.Value);
    }

    /// <summary>Adds an RRULE property.</summary>
    public static ContentLineProperty AddRecurrenceRule(this ContentLineComponent component,
        IcsRecurrenceRule rule) {
        if (component == null) throw new ArgumentNullException(nameof(component));
        if (rule == null) throw new ArgumentNullException(nameof(rule));
        return component.AddProperty("RRULE", rule.ToString());
    }
}
