namespace OfficeIMO.Html;

/// <summary>
/// Market-facing HTML scenario used to grow examples, galleries, and regression evidence.
/// </summary>
public sealed class HtmlMarketScenario {
    internal HtmlMarketScenario(string id, string title, HtmlConversionProfile profile, IEnumerable<string> capabilities, string customerValue) {
        Id = id ?? throw new ArgumentNullException(nameof(id));
        Title = title ?? throw new ArgumentNullException(nameof(title));
        Profile = profile;
        Capabilities = (capabilities ?? throw new ArgumentNullException(nameof(capabilities))).ToList().AsReadOnly();
        CustomerValue = customerValue ?? throw new ArgumentNullException(nameof(customerValue));
    }

    /// <summary>Stable scenario id.</summary>
    public string Id { get; }

    /// <summary>Display title.</summary>
    public string Title { get; }

    /// <summary>Recommended conversion profile.</summary>
    public HtmlConversionProfile Profile { get; }

    /// <summary>Capability tags the scenario must exercise.</summary>
    public IReadOnlyList<string> Capabilities { get; }

    /// <summary>Plain-language customer value.</summary>
    public string CustomerValue { get; }
}
