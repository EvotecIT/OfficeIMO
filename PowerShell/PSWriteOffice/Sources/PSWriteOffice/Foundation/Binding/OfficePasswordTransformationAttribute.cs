using System;
using System.Management.Automation;
using System.Security;

namespace PSWriteOffice.Foundation;

/// <summary>Binds a SecureString or a credential password; warns before accepting plain text.</summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
public sealed class OfficePasswordTransformationAttribute : ArgumentTransformationAttribute {
    /// <inheritdoc />
    public override object Transform(EngineIntrinsics engineIntrinsics, object inputData) => ConvertPassword(inputData,
        message => engineIntrinsics.InvokeCommand.InvokeScript("param($message) Microsoft.PowerShell.Utility\\Write-Warning -Message $message", false,
            System.Management.Automation.Runspaces.PipelineResultTypes.None, null, message));

    internal static SecureString ConvertPassword(object inputData, Action<string> warning) {
        var value = OfficeArgumentValue.Unwrap(inputData);
        if (value is SecureString secure) return secure;
        if (value is PSCredential credential) return credential.Password;
        if (value is string text) {
            warning("A plain-text password was supplied. Use Read-Host -AsSecureString or a PSCredential to keep passwords out of scripts and history.");
            var result = new SecureString();
            foreach (char character in text) result.AppendChar(character);
            result.MakeReadOnly();
            return result;
        }
        throw OfficeArgumentValue.Invalid("Expected a SecureString, PSCredential or plain-text password.");
    }
}
