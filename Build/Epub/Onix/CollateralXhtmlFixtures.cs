using OfficeIMO.Workflows;

internal static class CollateralXhtmlFixtures {
    internal static IReadOnlyList<BookOnixCollateralText> Create() => [new() {
        Type = BookOnixTextType.Description, Audiences = [BookOnixContentAudience.EndCustomers],
        Texts = [new("<div xmlns='http://www.w3.org/1999/xhtml' xml:lang='en'><h2>Rich description</h2>" +
            "<p>A <strong>clear</strong> &amp; <em>formatted</em> description with <a href='https://example.org/?a=1&amp;b=2'>a link</a>.</p>" +
            "<p><strong>First</strong> <em>second</em></p><pre>  <code>A</code>\n  <code>B</code>  </pre>" +
            "<ol start='3' type='i'><li>First item</li><li>Second item</li></ol>" +
            "<dl><dt>Term</dt><dd>Definition</dd></dl><blockquote cite='https://example.org/source'><p>Example quotation.</p></blockquote>" +
            "<table><caption>Details</caption><thead><tr><th scope='col'>Name</th><th scope='col'>Value</th></tr></thead>" +
            "<tbody><tr><td>Format</td><td>EPUB</td></tr></tbody></table><p><code>code</code><br/>Second line.</p></div>", "eng") {
                Format = BookOnixCollateralTextFormat.Xhtml }]
    }, new() { Type = BookOnixTextType.ShortDescription, Audiences = [BookOnixContentAudience.Unrestricted],
        Texts = [new("<p><strong>" + string.Concat(Enumerable.Repeat("&#x1F600;", 350)) + "</strong></p>", "eng") {
            Format = BookOnixCollateralTextFormat.Xhtml }] }
    ];
}
