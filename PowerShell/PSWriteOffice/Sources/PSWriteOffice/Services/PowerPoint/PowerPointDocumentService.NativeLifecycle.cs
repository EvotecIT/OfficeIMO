using System.Management.Automation;
using System.Security;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Services.PowerPoint;

public static partial class PowerPointDocumentService {
    /// <summary>Creates native save policy without resetting loaded encryption when Password was omitted.</summary>
    internal static PowerPointSaveOptions NativeSaveOptions(bool allowLoss, SecureString? password, bool passwordBound) {
        var options = new PowerPointSaveOptions { AllowLoss = allowLoss };
        if (passwordBound) options.Password = password.ToPlainText();
        return options;
    }

    /// <summary>Runs authoring on the pipeline thread inside the presentation's content scope.</summary>
    internal static void InvokeContent(PowerPointDocument document, ScriptBlock content) {
        using (PSWriteOffice.Foundation.OfficeDocumentContext.Push(document)) content.Invoke();
    }
}
