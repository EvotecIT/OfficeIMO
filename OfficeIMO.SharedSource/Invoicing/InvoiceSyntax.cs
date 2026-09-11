using System;
using System.IO;
using System.Linq;

namespace OfficeIMO.Internal.Invoicing;

/// <summary>Supported XML syntax families, independently of the declared invoice profile.</summary>
internal enum InvoiceSyntax {
    /// <summary>UN/CEFACT CrossIndustryInvoice namespace 100.</summary>
    Cii,
    /// <summary>OASIS UBL Invoice or CreditNote namespace 2.</summary>
    Ubl
}
