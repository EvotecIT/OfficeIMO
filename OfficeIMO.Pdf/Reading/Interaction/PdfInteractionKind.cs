namespace OfficeIMO.Pdf;

/// <summary>Kind of interactive or selectable page region.</summary>
public enum PdfInteractionKind {
    /// <summary>Approximate geometry for one extracted Unicode text element.</summary>
    Text = 0,

    /// <summary>Link annotation hit region.</summary>
    Link = 1,

    /// <summary>Non-link, non-widget annotation hit region.</summary>
    Annotation = 2,

    /// <summary>AcroForm widget hit region.</summary>
    FormWidget = 3,

    /// <summary>One exact image placement invocation on the page.</summary>
    Image = 4
}
