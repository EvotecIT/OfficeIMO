param(
    [ValidateSet('Manifest', 'Build', 'Publish')]
    [string] $RunMode = 'Build',

    [bool] $SignModule = $true,

    [string] $StagingPath,

    [switch] $SkipInstall,

    [switch] $PassThru
)

Import-Module PSPublishModule -MinimumVersion 3.0.168 -Force -ErrorAction Stop

Build-Module -ModuleName 'PSWriteOffice' -RunMode $RunMode {
    # Usual defaults as per standard module
    $Manifest = [ordered] @{
        # Minimum version of the Windows PowerShell engine required by this module
        PowerShellVersion      = '5.1'
        # prevent using over CORE/PS 7
        CompatiblePSEditions   = @('Desktop', 'Core')
        # ID used to uniquely identify this module
        GUID                   = 'd75a279d-30c2-4c2d-ae0d-12f1f3bf4d39'
        # Version number of this module.
        ModuleVersion          = '3.0.X'
        # Author of this module
        Author                 = 'Przemyslaw Klys'
        # Company or vendor of this module
        CompanyName            = 'Evotec'
        # Copyright statement for this module
        Copyright              = "(c) 2011 - $((Get-Date).Year) Przemyslaw Klys @ Evotec. All rights reserved."
        # Description of the functionality provided by this module
        Description            = 'PowerShell module to create, convert, read, inspect, and export Office, PDF, iWork, OneNote, OpenDocument, email, Markdown, AsciiDoc, LaTeX, CSV, Visio, and Confluence content without Microsoft Office installed. Powered by OfficeIMO.*.'
        # Tags applied to this module. These help with module discovery in online galleries.
        Tags                   = @('officeimo', 'word', 'excel', 'powerpoint', 'pdf', 'iwork', 'pages', 'numbers', 'keynote', 'onenote', 'ocr', 'provenance', 'security', 'opendocument', 'email', 'markdown', 'asciidoc', 'latex', 'csv', 'confluence', 'adf', 'atlassian', 'docx', 'xlsx', 'pptx', 'openxml', 'windows', 'linux', 'macos')
        # A URL to the main website for this project.
        ProjectUri             = 'https://github.com/EvotecIT/PSWriteOffice'

        LicenseUri             = 'https://github.com/EvotecIT/PSWriteOffice/blob/master/License'

        DotNetFrameworkVersion = '4.7.2'
    }
    New-ConfigurationManifest @Manifest

    $ConfigurationFormat = [ordered] @{
        RemoveComments                              = $false

        PlaceOpenBraceEnable                        = $true
        PlaceOpenBraceOnSameLine                    = $true
        PlaceOpenBraceNewLineAfter                  = $true
        PlaceOpenBraceIgnoreOneLineBlock            = $false

        PlaceCloseBraceEnable                       = $true
        PlaceCloseBraceNewLineAfter                 = $true
        PlaceCloseBraceIgnoreOneLineBlock           = $false
        PlaceCloseBraceNoEmptyLineBefore            = $true

        UseConsistentIndentationEnable              = $true
        UseConsistentIndentationKind                = 'space'
        UseConsistentIndentationPipelineIndentation = 'IncreaseIndentationAfterEveryPipeline'
        UseConsistentIndentationIndentationSize     = 4

        UseConsistentWhitespaceEnable               = $true
        UseConsistentWhitespaceCheckInnerBrace      = $true
        UseConsistentWhitespaceCheckOpenBrace       = $true
        UseConsistentWhitespaceCheckOpenParen       = $true
        UseConsistentWhitespaceCheckOperator        = $true
        UseConsistentWhitespaceCheckPipe            = $true
        UseConsistentWhitespaceCheckSeparator       = $true

        AlignAssignmentStatementEnable              = $true
        AlignAssignmentStatementCheckHashtable      = $true

        UseCorrectCasingEnable                      = $true
    }
    # format PSD1 and PSM1 files when merging into a single file
    # enable formatting is not required as Configuration is provided
    New-ConfigurationFormat -ApplyTo 'OnMergePSM1', 'OnMergePSD1' -Sort None @ConfigurationFormat
    # format PSD1 and PSM1 files within the module
    # enable formatting is required to make sure that formatting is applied (with default settings)
    New-ConfigurationFormat -ApplyTo 'DefaultPSD1', 'DefaultPSM1' -EnableFormatting -Sort None
    # when creating PSD1 use special style without comments and with only required parameters
    New-ConfigurationFormat -ApplyTo 'DefaultPSD1', 'OnMergePSD1' -PSD1Style 'Minimal'
    # configuration for documentation, at the same time it enables documentation processing
    New-ConfigurationDocumentation -Enable:$true -PathReadme 'Docs\Readme.md' -Path 'Docs'

    New-ConfigurationImportModule -ImportSelf

    $newConfigurationBuildSplat = @{
        Enable                            = $true
        # Release and local builds sign by default; CI smoke builds opt out explicitly.
        SignModule                        = $SignModule
        MergeModuleOnBuild                = $true
        MergeFunctionsFromApprovedModules = $true
        CertificateThumbprint             = '92e95fb58effa6a4a75e77a33cdd6bfe6dd30f1a'
        ResolveBinaryConflicts            = $true
        ResolveBinaryConflictsName        = 'PSWriteOffice'
        NETProjectName                    = 'PSWriteOffice'
        NETConfiguration                  = 'Release'
        NETFramework                      = 'net8.0', 'net472'
        NETHandleAssemblyWithSameName     = $true
        NETProjectPath                    = 'Sources\PSWriteOffice\PSWriteOffice.csproj'
        NETAssemblyLoadContext            = $true
        #NETMergeLibraryDebugging          = $true
        DotSourceLibraries                = $true
        DotSourceClasses                  = $true
        DeleteTargetModuleBeforeBuild     = $true
        NETBinaryModuleDocumentation      = $true
    }

    New-ConfigurationBuild @newConfigurationBuildSplat

    New-ConfigurationArtefact -Type Unpacked -Enable -Path 'Artefacts\Unpacked' -ModulesPath 'Modules' -RequiredModulesPath 'Modules' -AddRequiredModules -CopyFilesRelative
    New-ConfigurationArtefact -Type Packed -Enable -Path 'Artefacts\Packed' -ArtefactName '<ModuleName>.v<ModuleVersion>.zip'

    # global options for publishing to github/psgallery
    New-ConfigurationPublish -Type PowerShellGallery -FilePath 'C:\Support\Important\PowerShellGalleryAPI.txt' -Enabled:$true
    New-ConfigurationPublish -Type GitHub -FilePath 'C:\Support\Important\GitHubAPI.txt' -UserName 'EvotecIT' -Enabled:$true -GenerateReleaseNotes
} -StagingPath $StagingPath -SkipInstall:$SkipInstall -PassThru:$PassThru -ExitCode:(-not $PassThru)
