---
title: Protect and Encrypt Word Documents
description: "Restrict editing of DOCX files to read-only, comments, tracked changes or forms, mark them final, and encrypt or open them with a password in .NET."
order: 19
packages:
  - OfficeIMO.Word
cmdlets:
  - Protect-OfficeDocument
  - Unprotect-OfficeDocument
  - Open-OfficeWord
limits: Editing restrictions ask Word to limit editing and are not a security boundary; use encryption when the content must stay private.
search_terms:
  - password protect word
  - encrypt docx
  - read only document
  - restrict editing
  - mark as final
verified: 2026-11-14
---

Word has two kinds of protection, and they solve different problems. **Editing restrictions** ask Word to allow only reading, comments, tracked changes, or form filling; anyone with other tools can still read the content. **Encryption** scrambles the whole file with a password, so it cannot be opened without it. Apply either one last, after all content changes, and save to a new file.

## Restrict editing

`ProtectionType` selects what Word allows: `ReadOnly`, `Comments`, `TrackedChanges`, or `Forms`. With a password, Word asks for it before the restriction can be turned off.

### C#

```csharp
using OfficeIMO.Word;

using WordDocument document = WordDocument.Create("policy.docx");
document.AddParagraph("Acceptable use policy").Style = WordParagraphStyles.Heading1;
document.AddParagraph("This policy applies to every employee.");

document.Settings.ProtectionType = WordDocumentProtectionType.ReadOnly;
document.Settings.ProtectionPassword = Environment.GetEnvironmentVariable("POLICY_EDIT_PASSWORD")
    ?? throw new InvalidOperationException("Set POLICY_EDIT_PASSWORD.");
document.Save();
```

### PowerShell

```powershell
Protect-OfficeDocument -Path .\policy.docx -Restriction ReadOnly -Password (Read-Host -AsSecureString) -OutputPath .\policy-locked.docx
```

## Allow only tracked changes or comments

Reviewers can still edit, but every change is tracked, or they can only add comments.

### C#

```csharp
using OfficeIMO.Word;

using WordDocument document = WordDocument.Load("contract.docx");
document.Settings.ProtectionType = WordDocumentProtectionType.TrackedChanges;
document.Save("contract-for-review.docx");
```

### PowerShell

```powershell
Protect-OfficeDocument -Path .\contract.docx -Restriction TrackedChanges -OutputPath .\contract-for-review.docx
Protect-OfficeDocument -Path .\contract.docx -Restriction Comments -OutputPath .\contract-for-comments.docx
```

## Mark a document as final or recommend read-only

These are hints, not restrictions: Word shows a banner or asks whether to open read-only, and the reader can choose to edit anyway.

### C#

```csharp
using OfficeIMO.Word;

using WordDocument document = WordDocument.Load("minutes.docx");
document.Settings.FinalDocument = true;
document.Settings.ReadOnlyRecommended = true;
document.Save("minutes-final.docx");
```

## Encrypt a document with a password

`WordSaveOptions.Password` encrypts the saved file. Setting it to `null` explicitly removes encryption from a document that was loaded encrypted.

### C#

```csharp
using OfficeIMO.Word;

string password = Environment.GetEnvironmentVariable("PAYROLL_PASSWORD")
    ?? throw new InvalidOperationException("Set PAYROLL_PASSWORD.");

using WordDocument document = WordDocument.Create();
document.AddParagraph("Payroll summary");
document.Save("payroll.docx", new WordSaveOptions { Password = password });
```

### PowerShell

```powershell
Get-ChildItem .\Payroll -Filter *.docx | Protect-OfficeDocument -Encrypt -Password (Read-Host -AsSecureString)
```

## Open an encrypted document

### C#

```csharp
using OfficeIMO.Word;

string password = Environment.GetEnvironmentVariable("PAYROLL_PASSWORD")
    ?? throw new InvalidOperationException("Set PAYROLL_PASSWORD.");

using WordDocument document = WordDocument.Load("payroll.docx", new WordLoadOptions { Password = password });
Console.WriteLine(document.Paragraphs[0].Text);

document.Save("payroll-open.docx", new WordSaveOptions { Password = null });
```

### PowerShell

```powershell
$password = Read-Host -AsSecureString
$document = Open-OfficeWord -Path .\payroll.docx -Password $password
Get-OfficeWordText -Document $document
Close-OfficeWord -Document $document

Unprotect-OfficeDocument -Path .\payroll.docx -Password $password -OutputPath .\payroll-open.docx
```
