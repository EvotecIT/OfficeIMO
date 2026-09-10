using System;
using System.IO;
using System.Linq;

namespace OfficeIMO.Internal.Invoicing;

/// <summary>Declared electronic invoice profiles. Recognition alone does not establish compliance.</summary>
internal enum InvoiceProfile {
    /// <summary>Factur-X header subset.</summary>
    Minimum,
    /// <summary>Factur-X basic profile without invoice lines.</summary>
    BasicWithoutLines,
    /// <summary>Factur-X basic profile including core line information.</summary>
    Basic,
    /// <summary>The EN 16931 semantic invoice model.</summary>
    En16931,
    /// <summary>Factur-X extended profile.</summary>
    Extended,
    /// <summary>German XRechnung 3.0 usage specification.</summary>
    XRechnung,
    /// <summary>Peppol BIS Billing 3.0 usage specification.</summary>
    PeppolBis,
    /// <summary>French extended invoice profile. Recognition does not imply authoring or rule coverage.</summary>
    ExtendedCtcFr
}
