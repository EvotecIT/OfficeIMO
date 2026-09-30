using System;

namespace OfficeIMO.Internal.Invoicing;

internal sealed class InvoiceSourceFieldException : InvalidOperationException {
    internal InvoiceSourceFieldException(InvoiceSourceField field, Exception failure) : base(failure.Message, failure) => Field = field;
    internal InvoiceSourceField Field { get; }
}
