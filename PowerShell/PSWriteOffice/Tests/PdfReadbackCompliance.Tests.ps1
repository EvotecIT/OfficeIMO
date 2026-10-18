BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop
}

Describe 'PDF readback and compliance cmdlets' {
    It 'converts PDF logical readback to Markdown' {
        $pdfPath = Join-Path $TestDrive 'markdown-source.pdf'
        $markdownPath = Join-Path $TestDrive 'markdown-output.md'

        New-OfficePdf -Path $pdfPath {
            PdfHeading 'Markdown Source'
            PdfParagraph 'Logical readback should preserve this sentence.'
        } | Out-Null

        ConvertTo-OfficePdfMarkdown -Path $pdfPath -OutputPath $markdownPath | Should -BeOfType System.IO.FileInfo

        $markdown = Get-Content -Raw -Path $markdownPath
        $markdown | Should -Match 'Markdown Source'
        $markdown | Should -Match 'Logical readback'
    }

    It 'extracts image resources from generated PDFs' {
        $pdfPath = Join-Path $TestDrive 'images.pdf'
        $imagePath = Join-Path $PSScriptRoot 'Assets\CellImage.png'
        $outputDirectory = Join-Path $TestDrive 'images'

        New-OfficePdf -Path $pdfPath {
            PdfHeading 'Image Report'
            PdfImage -Path $imagePath -Width 64 -Height 64 -AlternativeText 'Sample image'
        } | Out-Null

        $images = @(Get-OfficePdfImage -Path $pdfPath)
        $images.Count | Should -BeGreaterThan 0
        $images[0].Width | Should -BeGreaterThan 0

        $files = @(Get-OfficePdfImage -Path $pdfPath -OutputDirectory $outputDirectory -BaseName 'asset:review?')
        $files.Count | Should -BeGreaterThan 0
        Test-Path $files[0].FullName | Should -BeTrue
        $files[0].Name | Should -Match '^asset_review_-0001\.'
    }

    It 'embeds and extracts PDF attachments' {
        $pdfPath = Join-Path $TestDrive 'attachments.pdf'
        $attachmentPath = Join-Path $TestDrive 'payload.txt'
        $outputDirectory = Join-Path $TestDrive 'attachments'
        Set-Content -Path $attachmentPath -Value 'Attachment payload' -NoNewline

        New-OfficePdf -Path $pdfPath {
            PdfHeading 'Attachment Report'
            PdfAttachment -Path $attachmentPath -Name 'payload.txt' -MimeType 'text/plain' -Relationship Data -Description 'Test payload'
        } | Out-Null

        $attachments = @(Get-OfficePdfAttachment -Path $pdfPath)
        $attachments.Count | Should -Be 1
        $attachments[0].FileName | Should -Be 'payload.txt'
        $attachments[0].Relationship | Should -Be 'Data'

        $files = @(Get-OfficePdfAttachment -Path $pdfPath -OutputDirectory $outputDirectory)
        $files.Count | Should -Be 1
        Get-Content -Raw -Path $files[0].FullName | Should -Be 'Attachment payload'
    }

    It 'normalizes untrusted extracted file names with portable rules' {
        $assembly = (Get-Command Get-OfficePdfAttachment).ImplementingType.Assembly
        $utilities = $assembly.GetType('PSWriteOffice.Services.Pdf.PdfCommandUtilities', $true)
        $method = $utilities.GetMethod(
            'GetSafeFileName',
            [Reflection.BindingFlags]::Static -bor [Reflection.BindingFlags]::NonPublic)

        $method.Invoke($null, @('..\folder/payload:report?.txt')) | Should -Be 'payload_report_.txt'
        $method.Invoke($null, @('CON.txt')) | Should -Be '_CON.txt'
        $method.Invoke($null, @('CON.tar.gz')) | Should -Be '_CON.tar.gz'
        $method.Invoke($null, @('LPT1.backup.txt')) | Should -Be '_LPT1.backup.txt'
    }

    It 'keeps context attachments, metadata, and encryption in one generated document' {
        $pdfPath = Join-Path $TestDrive 'encrypted-attachment.pdf'
        $attachmentPath = Join-Path $TestDrive 'encrypted-payload.txt'
        Set-Content -Path $attachmentPath -Value 'Encrypted attachment payload' -NoNewline

        New-OfficePdf -Path $pdfPath -Password 'open' {
            PdfAttachment -Path $attachmentPath -Name 'encrypted-payload.txt' -MimeType 'text/plain'
            PdfMetadata -Title 'Encrypted attachment report'
            PdfParagraph 'Encrypted document content'
        } | Out-Null

        { Get-OfficePdfAttachment -Path $pdfPath } | Should -Throw
        (Get-OfficePdfPreflight -Path $pdfPath -Password 'open').CanRead | Should -BeTrue
        (Get-OfficePdfAttachment -Path $pdfPath -Password 'open').FileName | Should -Be 'encrypted-payload.txt'
        (Get-OfficePdfInfo -Path $pdfPath -Password 'open').Metadata.Title | Should -Be 'Encrypted attachment report'
    }

    It 'always returns the replacement document for explicit attachment rewrites' {
        $pdfPath = Join-Path $TestDrive 'explicit-attachment.pdf'
        $attachmentPath = Join-Path $TestDrive 'explicit-payload.txt'
        Set-Content -Path $attachmentPath -Value 'Explicit attachment payload' -NoNewline
        $document = New-OfficePdf {
            PdfParagraph 'Explicit attachment document'
        }

        $updated = $document | PdfAttachment -Path $attachmentPath -Name 'explicit-payload.txt' -PassThru
        $updated | Save-OfficePdf -Path $pdfPath | Out-Null

        $updated | Should -BeOfType ($document.GetType())
        (Get-OfficePdfAttachment -Path $pdfPath).FileName | Should -Be 'explicit-payload.txt'
    }

    It 'reports generated PDF compliance readiness' {
        $document = New-OfficePdf {
            PdfMetadata -Title 'Compliance Draft'
            PdfCompliance -Profile PdfUa1 -Groundwork
            PdfHeading 'Compliance Draft'
            PdfParagraph 'This document can be assessed before saving.'
        }

        $report = $document | Get-OfficePdfCompliance -Profile PdfUa1
        $report.Profile | Should -Be 'PdfUa1'
        $report.DisplayName | Should -BeLike '*PDF/UA*'
        $report.Requirements.Count | Should -BeGreaterThan 0
    }

    It 'requires artifact-bound external validator evidence for compliance claims' {
        $document = New-OfficePdf {
            PdfMetadata -Title 'Compliance Proof Draft'
            PdfCompliance -Profile PdfA3B -Groundwork
            PdfParagraph 'Compliance proof combines readiness with external validation evidence.'
        }

        $proof = $document | Get-OfficePdfCompliance -Profile PdfA3B -Proof

        $proof.Profile | Should -Be 'PdfA3B'
        $proof.RequiredExternalValidatorCount | Should -BeGreaterThan 0
        $proof.RequiresExternalValidation | Should -BeTrue
        $proof.MissingExternalValidatorCount | Should -BeGreaterThan 0
        $proof.ProofStatus | Should -Be 'InternalGaps'
        $proof.CanClaimConformance | Should -BeFalse

        $unbound = $document | Get-OfficePdfCompliance -Profile PdfA3B -Proof -ExternalValidator VeraPdf -ExternalStatus Passed -ExternalValidatorVersion '1.26.1' -ExternalProfile 'PDF/A-3b' -ExternalDiagnostic 'caller reported veraPDF passed'
        $unbound.PassedExternalValidationCount | Should -Be 0
        $unbound.ExternalValidations[0].HasArtifactBinding | Should -BeFalse
        $unbound.HasRequiredExternalValidation | Should -BeFalse
        $unbound.MissingExternalValidatorCount | Should -BeGreaterThan 0
        $unbound.CanClaimConformance | Should -BeFalse

        $artifact = $document.CreateComplianceArtifact($proof.Profile)
        $validationType = $unbound.ExternalValidations[0].GetType()
        $validation = $validationType.GetMethod('PassedForArtifact').Invoke($null, [object[]] @(
            $unbound.ExternalValidations[0].ValidatorKind,
            'veraPDF',
            '1.26.1',
            'veraPDF passed',
            $artifact.ToBytes(),
            'PDF/A-3b',
            $null,
            $null,
            $null,
            0,
            $null))
        $withEvidence = $document | Get-OfficePdfCompliance -Profile PdfA3B -Proof -ExternalValidation $validation
        $withEvidence.PassedExternalValidationCount | Should -Be 1
        $withEvidence.HasArtifactEvidence | Should -BeTrue
        $withEvidence.ExternalValidations[0].HasArtifactBinding | Should -BeTrue
        $withEvidence.ExternalValidations[0].ArtifactSha256 | Should -Be $withEvidence.ArtifactSha256
        $withEvidence.ExternalValidations[0].ValidatorVersion | Should -Be '1.26.1'
        $withEvidence.ExternalValidations[0].Diagnostic | Should -Be 'veraPDF passed'

        $fromExitCode = $document | Get-OfficePdfCompliance -Profile PdfA3B -Proof -ExternalValidator VeraPdf -ExternalProfile 'PDF/A-3b' -ExternalDiagnostic 'veraPDF process result' -ExternalExitCode 0 -ExternalExecutablePath 'verapdf' -ExternalArguments '--format text file.pdf'
        $fromExitCode.PassedExternalValidationCount | Should -Be 0
        $fromExitCode.ExternalValidations[0].HasArtifactBinding | Should -BeFalse
        $fromExitCode.HasRequiredExternalValidation | Should -BeFalse
        $fromExitCode.ExternalValidations[0].ExitCode | Should -Be 0
        $fromExitCode.ExternalValidations[0].ExecutablePath | Should -Be 'verapdf'
        $fromExitCode.ExternalValidations[0].Arguments | Should -Be '--format text file.pdf'
    }

    It 'gets font diagnostics as pipeline-friendly objects' {
        (Get-Command Get-OfficePdfFont).Parameters.Keys | Should -Contain 'NeedsReview'
        $pdfPath = Join-Path $TestDrive 'font-diagnostics.pdf'

        New-OfficePdf -Path $pdfPath {
            PdfParagraph 'Font diagnostic readback'
        } | Out-Null

        $fonts = @(Get-OfficePdfFont -Path $pdfPath)
        $fonts.Count | Should -BeGreaterThan 0
        $fonts[0].PSObject.Properties.Name | Should -Contain 'RepairReadiness'
        $fonts[0].PSObject.Properties.Name | Should -Contain 'HasToUnicodeMap'
    }

    It 'gets annotation readback with action filters' {
        (Get-Command Get-OfficePdfAnnotation).Parameters.Keys | Should -Contain 'WithAction'
        Get-Command Set-OfficePdfAnnotation | Should -Not -BeNullOrEmpty
        Get-Command Remove-OfficePdfAnnotation | Should -Not -BeNullOrEmpty
        $pdfPath = Join-Path $TestDrive 'annotations.pdf'
        $updatedPath = Join-Path $TestDrive 'annotations-updated.pdf'
        $removedPath = Join-Path $TestDrive 'annotations-removed.pdf'
        @'
%PDF-1.7
1 0 obj
<< /Type /Catalog /Pages 2 0 R >>
endobj
2 0 obj
<< /Type /Pages /Count 1 /Kids [3 0 R] >>
endobj
3 0 obj
<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Annots [4 0 R] >>
endobj
4 0 obj
<< /Type /Annot /Subtype /Text /Rect [20 20 40 40] /Contents (Review note) /AA << /E << /S /Launch /F (tool.exe) >> >> >>
endobj
trailer
<< /Root 1 0 R /Size 5 >>
startxref
123
%%EOF
'@ | Set-Content -Path $pdfPath -NoNewline -Encoding Ascii

        $annotations = @(Get-OfficePdfAnnotation -Path $pdfPath -Subtype Text -WithAction)
        $annotations.Count | Should -Be 1
        $annotations[0].Contents | Should -Be 'Review note'
        $annotations[0].HasAdditionalActions | Should -BeTrue

        @'
%PDF-1.7
1 0 obj
<< /Type /Catalog /Pages 2 0 R >>
endobj
2 0 obj
<< /Type /Pages /Count 1 /Kids [3 0 R] >>
endobj
3 0 obj
<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Annots [4 0 R] >>
endobj
4 0 obj
<< /Type /Annot /Subtype /Text /Rect [20 20 40 40] /Contents (Editable note) >>
endobj
trailer
<< /Root 1 0 R /Size 5 >>
startxref
123
%%EOF
'@ | Set-Content -Path $pdfPath -NoNewline -Encoding Ascii

        $updateReport = Set-OfficePdfAnnotation -Path $pdfPath -OutputPath $updatedPath -ObjectNumber 4 -Contents 'Updated note' -Title 'Reviewer' -Name 'Note-2' -Color '#0080FF' -RemoveAction -PassThruReport -PassThru
        $updateReport.Applied | Should -BeTrue
        $updated = @(Get-OfficePdfAnnotation -Path $updatedPath -Subtype Text)
        $updated.Count | Should -Be 1
        $updated[0].Contents | Should -Be 'Updated note'
        $updated[0].Title | Should -Be 'Reviewer'
        $updated[0].Name | Should -Be 'Note-2'
        $updated[0].HasAdditionalActions | Should -BeFalse

        $removeReport = Remove-OfficePdfAnnotation -Path $updatedPath -OutputPath $removedPath -Subtype Text -PassThruReport -PassThru
        $removeReport.Applied | Should -BeTrue
        @(Get-OfficePdfAnnotation -Path $removedPath -Subtype Text).Count | Should -Be 0
    }

    It 'configures PDF/A-4 and PDF/UA-2 groundwork through compliance profiles' {
        $pdfA4Path = Join-Path $TestDrive 'pdfa4.pdf'
        $pdfUa2Path = Join-Path $TestDrive 'pdfua2.pdf'

        New-OfficePdf -Path $pdfA4Path {
            PdfCompliance -Profile PdfA4F -Groundwork -Language 'en-US'
            PdfMetadata -Title 'PDF/A-4F Draft'
            PdfParagraph 'PDF/A-4F groundwork from PowerShell.'
        } | Out-Null

        New-OfficePdf -Path $pdfUa2Path {
            PdfCompliance -Profile PdfUa2 -Groundwork -Language 'en-US'
            PdfMetadata -Title 'PDF/UA-2 Draft'
            PdfHeading 'PDF/UA-2 Draft'
            PdfParagraph 'PDF/UA-2 groundwork from PowerShell.'
        } | Out-Null

        $pdfA4Info = Get-OfficePdfInfo -Path $pdfA4Path
        $pdfUa2Info = Get-OfficePdfInfo -Path $pdfUa2Path
        $pdfA4Report = Get-OfficePdfCompliance -Path $pdfA4Path -Profile PdfA4F
        $pdfUa2Report = Get-OfficePdfCompliance -Path $pdfUa2Path -Profile PdfUa2

        $pdfA4Info.HeaderVersion | Should -Be '2.0'
        $pdfA4Info.IsPdf20OrLater | Should -BeTrue
        $pdfA4Report.FindRequirement('readback-pdfa-identification').Status | Should -Be 'Satisfied'
        $pdfA4Report.FindRequirement('readback-output-intent').Status | Should -Be 'Satisfied'

        $pdfUa2Info.HeaderVersion | Should -Be '2.0'
        $pdfUa2Info.TaggedContent.HasDocumentStructureElement | Should -BeTrue
        $pdfUa2Info.TaggedContent.MarkedContentReferenceCount | Should -BeGreaterThan 0
        $pdfUa2Report.FindRequirement('readback-pdfua-identification').Status | Should -Be 'Satisfied'
        $pdfUa2Report.FindRequirement('readback-structure-element-count').Status | Should -Be 'Satisfied'
        $pdfUa2Report.FindRequirement('readback-marked-content-references').Status | Should -Be 'Satisfied'
        $pdfUa2Report.FindRequirement('pdfua-validation').Status | Should -Be 'Unsupported'
    }

    It 'reports saved PDF compliance readback evidence by path' {
        $pdfPath = Join-Path $TestDrive 'pdfua-readback.pdf'

        New-OfficePdf -Path $pdfPath {
            PdfCompliance -Profile PdfUa1 -Groundwork -Language 'en-US'
            PdfMetadata -Title 'Readback PDF/UA' -Author 'PSWriteOffice'
            PdfHeading 'Readback PDF/UA'
            PdfParagraph 'Saved PDF compliance readback evidence'
        } | Out-Null

        $report = Get-OfficePdfCompliance -Path $pdfPath -Profile PdfUa1

        $report.Profile | Should -Be 'PdfUa1'
        $report.FindRequirement('readback-pdfua-identification').Status | Should -Be 'Satisfied'
        $report.FindRequirement('readback-document-title').Status | Should -Be 'Satisfied'
        $report.FindRequirement('readback-marked-catalog').Status | Should -Be 'Satisfied'
        $report.FindRequirement('pdfua-validation').Status | Should -Be 'Unsupported'
    }

    It 'configures e-invoice associated file groundwork from XML' {
        $xmlPath = Join-Path $TestDrive 'invoice.xml'
        $pdfPath = Join-Path $TestDrive 'einvoice.pdf'
        @'
<?xml version="1.0" encoding="UTF-8"?>
<rsm:CrossIndustryInvoice xmlns:rsm="urn:un:unece:uncefact:data:standard:CrossIndustryInvoice:100" xmlns:ram="urn:un:unece:uncefact:data:standard:ReusableAggregateBusinessInformationEntity:100">
  <rsm:ExchangedDocumentContext>
    <ram:GuidelineSpecifiedDocumentContextParameter>
      <ram:ID>urn:cen.eu:en16931:2017</ram:ID>
    </ram:GuidelineSpecifiedDocumentContextParameter>
  </rsm:ExchangedDocumentContext>
</rsm:CrossIndustryInvoice>
'@ | Set-Content -Path $xmlPath -Encoding UTF8

        New-OfficePdf -Path $pdfPath {
            PdfElectronicInvoice -Path $xmlPath -Profile FacturX -ConformanceLevel 'EN 16931'
            PdfMetadata -Title 'E-invoice'
            PdfHeading 'E-invoice'
            PdfParagraph 'The CII XML payload is embedded as an associated file.'
        } | Out-Null

        $attachment = Get-OfficePdfAttachment -Path $pdfPath
        $report = Get-OfficePdfCompliance -Path $pdfPath -Profile FacturX

        $attachment.FileName | Should -Be 'factur-x.xml'
        $attachment.MimeType | Should -Be 'application/xml'
        $attachment.Relationship | Should -Be 'Data'
        $report.FindRequirement('readback-einvoice-xmp').Status | Should -Be 'Satisfied'
        $report.FindRequirement('readback-associated-invoice-file').Status | Should -Be 'Satisfied'
    }
}
