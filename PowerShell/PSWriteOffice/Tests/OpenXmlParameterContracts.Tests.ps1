BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop

    function Get-ParameterLeafType {
        param([Parameter(Mandatory)][Type] $Type)

        $nullableType = [Nullable]::GetUnderlyingType($Type)
        if ($null -ne $nullableType) {
            Get-ParameterLeafType -Type $nullableType
            return
        }

        if ($Type.HasElementType) {
            Get-ParameterLeafType -Type ($Type.GetElementType())
            return
        }

        if ($Type.IsGenericType) {
            foreach ($argumentType in $Type.GetGenericArguments()) {
                Get-ParameterLeafType -Type $argumentType
            }
            return
        }

        $Type
    }

    function Get-PublicApiLeafType {
        param([Parameter(Mandatory)][Type] $Type)

        if ($Type.IsGenericParameter) {
            foreach ($constraintType in $Type.GetGenericParameterConstraints()) {
                Get-PublicApiLeafType -Type $constraintType
            }
            return
        }

        $nullableType = [Nullable]::GetUnderlyingType($Type)
        if ($null -ne $nullableType) {
            Get-PublicApiLeafType -Type $nullableType
            return
        }

        if ($Type.HasElementType) {
            Get-PublicApiLeafType -Type ($Type.GetElementType())
            return
        }

        if ($Type.IsGenericType) {
            foreach ($argumentType in $Type.GetGenericArguments()) {
                Get-PublicApiLeafType -Type $argumentType
            }
        }

        $Type
    }

    function Test-IsOpenXmlValueStruct {
        param([Parameter(Mandatory)][Type] $Type)

        $Type.IsValueType -and
            -not $Type.IsEnum -and
            @($Type.GetInterfaces()).FullName -contains 'DocumentFormat.OpenXml.IEnumValue'
    }

    function Get-PublicApiTypeUsage {
        param([Parameter(Mandatory)][Type] $DeclaringType)

        $usages = [System.Collections.Generic.List[object]]::new()
        $bindingFlags = [Reflection.BindingFlags]'Public,Instance,Static,DeclaredOnly'

        $addUsage = {
            param([string] $Location, [Type] $Type)
            foreach ($leafType in Get-PublicApiLeafType -Type $Type) {
                $usages.Add([pscustomobject]@{
                    Location = $Location
                    Type = $leafType
                })
            }
        }

        if ($null -ne $DeclaringType.BaseType) {
            & $addUsage "$($DeclaringType.FullName):base" $DeclaringType.BaseType
        }
        foreach ($interfaceType in $DeclaringType.GetInterfaces()) {
            & $addUsage "$($DeclaringType.FullName):interface" $interfaceType
        }
        foreach ($genericParameter in $DeclaringType.GetGenericArguments() | Where-Object IsGenericParameter) {
            & $addUsage "$($DeclaringType.FullName):generic-constraint:$($genericParameter.Name)" $genericParameter
        }

        foreach ($constructor in $DeclaringType.GetConstructors($bindingFlags)) {
            foreach ($parameter in $constructor.GetParameters()) {
                & $addUsage "$($DeclaringType.FullName):ctor:$($parameter.Name)" $parameter.ParameterType
            }
        }

        foreach ($method in $DeclaringType.GetMethods($bindingFlags)) {
            & $addUsage "$($DeclaringType.FullName):method:$($method.Name):return" $method.ReturnType
            foreach ($parameter in $method.GetParameters()) {
                & $addUsage "$($DeclaringType.FullName):method:$($method.Name):$($parameter.Name)" $parameter.ParameterType
            }
            foreach ($genericParameter in $method.GetGenericArguments() | Where-Object IsGenericParameter) {
                & $addUsage "$($DeclaringType.FullName):method:$($method.Name):generic-constraint:$($genericParameter.Name)" $genericParameter
            }
        }

        foreach ($property in $DeclaringType.GetProperties($bindingFlags)) {
            & $addUsage "$($DeclaringType.FullName):property:$($property.Name)" $property.PropertyType
            foreach ($parameter in $property.GetIndexParameters()) {
                & $addUsage "$($DeclaringType.FullName):property:$($property.Name):index:$($parameter.Name)" $parameter.ParameterType
            }
        }

        foreach ($field in $DeclaringType.GetFields($bindingFlags)) {
            & $addUsage "$($DeclaringType.FullName):field:$($field.Name)" $field.FieldType
        }

        foreach ($eventInfo in $DeclaringType.GetEvents($bindingFlags)) {
            if ($null -ne $eventInfo.EventHandlerType) {
                & $addUsage "$($DeclaringType.FullName):event:$($eventInfo.Name)" $eventInfo.EventHandlerType
            }
        }

        $usages
    }
}

Describe 'Open XML parameter contracts' {
    It 'does not expose Open XML SDK value structs as cmdlet parameters' {
        $leaks = Get-Command -Module PSWriteOffice | ForEach-Object {
            $command = $_
            $_.Parameters.GetEnumerator() | ForEach-Object {
                $parameter = $_
                Get-ParameterLeafType -Type $parameter.Value.ParameterType |
                    Where-Object { $_.FullName -like 'DocumentFormat.OpenXml*' } |
                    ForEach-Object { "$($command.Name):$($parameter.Key):$($_.FullName)" }
            }
        }

        @($leaks) | Should -BeNullOrEmpty
    }

    It 'does not expose Open XML SDK value structs anywhere in the public assembly API' {
        $assembly = [AppDomain]::CurrentDomain.GetAssemblies() |
            Where-Object { $_.GetName().Name -eq 'PSWriteOffice' } |
            Select-Object -First 1

        $assembly | Should -Not -BeNullOrEmpty

        $leaks = foreach ($publicType in $assembly.GetExportedTypes()) {
            Get-PublicApiTypeUsage -DeclaringType $publicType |
                Where-Object { Test-IsOpenXmlValueStruct -Type $_.Type } |
                ForEach-Object { "$($_.Location):$($_.Type.FullName)" }
        }

        @($leaks) | Should -BeNullOrEmpty
    }

    It 'uses OfficeIMO-owned CLR enums for PowerShell-facing document values' {
        $contracts = @{
            'Add-OfficeWordSection:BreakType' = 'OfficeIMO.Word.WordSectionBreakType'
            'Add-OfficeWordBreak:BreakType' = 'OfficeIMO.Word.WordBreakType'
            'Add-OfficeWordHeader:Type' = 'OfficeIMO.Word.WordHeaderFooterType'
            'Add-OfficeWordFooter:Type' = 'OfficeIMO.Word.WordHeaderFooterType'
            'Add-OfficeWordParagraph:Alignment' = 'OfficeIMO.Word.WordParagraphAlignment'
            'Add-OfficeWordText:Underline' = 'OfficeIMO.Word.WordUnderlineStyle'
            'New-OfficeWordTableCell:UnderlineStyle' = 'OfficeIMO.Word.WordUnderlineStyle'
            'New-OfficeWordTableCell:Align' = 'OfficeIMO.Word.WordParagraphAlignment'
            'New-OfficeWordTableCell:VerticalAlign' = 'OfficeIMO.Word.WordTableVerticalAlignment'
            'Add-OfficeWordTextBox:HorizontalPositionRelativeFrom' = 'OfficeIMO.Word.WordHorizontalRelativePosition'
            'Add-OfficeWordTextBox:VerticalPositionRelativeFrom' = 'OfficeIMO.Word.WordVerticalRelativePosition'
            'Protect-OfficeWordDocument:ProtectionType' = 'OfficeIMO.Word.WordDocumentProtectionType'
            'Set-OfficeWordTableCell:TextDirection' = 'OfficeIMO.Word.WordTextDirection'
            'Add-OfficePowerPointSlide:LayoutType' = 'OfficeIMO.PowerPoint.PowerPointSlideLayoutType'
            'Set-OfficePowerPointSlideLayout:LayoutType' = 'OfficeIMO.PowerPoint.PowerPointSlideLayoutType'
        }

        foreach ($entry in $contracts.GetEnumerator()) {
            $commandName, $parameterName = $entry.Key -split ':', 2
            $parameterType = (Get-Command $commandName).Parameters[$parameterName].ParameterType
            $nullableType = [Nullable]::GetUnderlyingType($parameterType)
            if ($null -ne $nullableType) {
                $parameterType = $nullableType
            }
            $parameterType.FullName | Should -Be $entry.Value
            $parameterType.IsEnum | Should -BeTrue
        }
    }
}
