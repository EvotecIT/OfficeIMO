using System;
using System.Management.Automation;
using OfficeIMO.Word;

namespace PSWriteOffice.Services.Word;

public static partial class WordDocumentService
{
    /// <summary>
    /// Executes a PowerShell Word DSL script block against an existing OfficeIMO document.
    /// </summary>
    public static void InvokeDsl(WordDocument document, ScriptBlock content)
    {
        if (document == null)
        {
            throw new ArgumentNullException(nameof(document));
        }

        if (content == null)
        {
            throw new ArgumentNullException(nameof(content));
        }

        using (WordDslContext.Enter(document))
        {
            content.InvokeReturnAsIs();
        }
    }
}
