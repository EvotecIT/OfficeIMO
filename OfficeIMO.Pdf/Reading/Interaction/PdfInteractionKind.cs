namespace OfficeIMO.Pdf;

/// <summary>Kind of interactive or selectable page region.</summary>
public enum PdfInteractionKind {
    /// <summary>Approximate geometry for one extracted Unicode text element.</summary>
    Text,

    /// <summary>Link annotation hit region.</summary>
    Link,

    /// <summary>Non-link, non-widget annotation hit region.</summary>
    Annotation,

    /// <summary>AcroForm widget hit region.</summary>
    FormWidget
}
