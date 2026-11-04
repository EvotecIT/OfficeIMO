using System;
using System.Collections;
using System.Collections.Generic;
using System.Management.Automation;
using System.Net;
using System.Security;

namespace PSWriteOffice.Foundation;

/// <summary>Binds an optional SecureString or credential password; warns before accepting plain text.</summary>
/// <para>A collection binds element by element for per-source password parameters; null elements stay null.</para>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
public sealed class OfficePasswordTransformationAttribute : ArgumentTransformationAttribute {
    private const string PlainTextWarning = "A plain-text password was supplied. Use Read-Host -AsSecureString or a PSCredential to keep passwords out of scripts and history.";

    /// <inheritdoc />
    public override object Transform(EngineIntrinsics engineIntrinsics, object inputData) {
        // The command decides whether null means an omitted source password or an explicit output reset.
        object? value = OfficeArgumentValue.Unwrap(inputData);
        if (value == null) return null!;
        bool warned = false;
        Action<string> warning = message => {
            if (warned) return;
            warned = true;
            engineIntrinsics.InvokeCommand.InvokeScript("param($message) Microsoft.PowerShell.Utility\\Write-Warning -Message $message", false,
                System.Management.Automation.Runspaces.PipelineResultTypes.None, null, message);
        };
        return value is IEnumerable values && value is not string ? ConvertPasswords(values, warning) : ConvertPassword(value, warning);
    }

    internal static SecureString ConvertPassword(object inputData, Action<string> warning) {
        var value = OfficeArgumentValue.Unwrap(inputData);
        if (value is SecureString secure) return secure;
        if (value is PSCredential credential) return credential.Password;
        if (value is string text) {
            warning(PlainTextWarning);
            var result = new SecureString();
            foreach (char character in text) result.AppendChar(character);
            result.MakeReadOnly();
            return result;
        }
        throw OfficeArgumentValue.Invalid("Expected a SecureString, PSCredential or plain-text password.");
    }

    internal static SecureString?[] ConvertPasswords(IEnumerable values, Action<string> warning) {
        var result = new List<SecureString?>();
        foreach (object? item in values)
            result.Add(OfficeArgumentValue.Unwrap(item) == null ? null : ConvertPassword(item!, warning));
        return result.ToArray();
    }
}

internal static class OfficePasswordExtensions {
    /// <summary>Reveals a bound password only at the native API boundary that requires a string.</summary>
    internal static string? ToPlainText(this SecureString? value) =>
        value == null ? null : new NetworkCredential(string.Empty, value).Password;
}
