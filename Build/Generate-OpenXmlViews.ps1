param(
    # Folder with the ISO/IEC 29500-4 transitional schemas (pml.xsd, dml-main.xsd, shared-*.xsd and the rest).
    [Parameter(Mandatory)] [string] $SchemaDirectory,
    # Complex types to generate views for, as '<prefix>:<type name>' with prefixes p, a, r, s or w.
    [Parameter(Mandatory)] [string[]] $Types,
    [Parameter(Mandatory)] [string] $Namespace,
    [Parameter(Mandatory)] [string] $OutputPath,
    [Parameter(Mandatory)] [string] $GeneratorName,
    # Groups, as '<prefix>:<group name>', whose choice the schema leaves unordered but Office writes in declaration
    # order, such as WordprocessingML run properties. Their alternatives get one rank each.
    [string[]] $OrderedChoiceGroups = @(),
    # A node type to build the views over in place of a tree and a node index, such as a handle that reaches both
    # OfficeIMO's trees and another XML layer. It has the members the generated code calls on it.
    [string] $NodeType = '',
    # The namespace that holds -NodeType, when the views are in another one.
    [string] $NodeNamespace = ''
)

# Writes typed views over OfficeXmlTree (or -NodeType) for schema complex types: attribute properties with schema value types, and
# child accessors that insert in the type's sequence order. Children whose types have no view still count for order.

$ErrorActionPreference = 'Stop'

$namespaces = [ordered]@{
    'p' = @{ Uri = 'http://schemas.openxmlformats.org/presentationml/2006/main'; Id = 'OfficeXmlNamespaces.PresentationML'; Prefix = 'P' }
    'a' = @{ Uri = 'http://schemas.openxmlformats.org/drawingml/2006/main'; Id = 'OfficeXmlNamespaces.DrawingML'; Prefix = 'A' }
    'r' = @{ Uri = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'; Id = 'OfficeXmlNamespaces.Relationships'; Prefix = 'R' }
    's' = @{ Uri = 'http://schemas.openxmlformats.org/officeDocument/2006/sharedTypes'; Id = 'OfficeXmlNamespaces.SharedTypes'; Prefix = 'S' }
    'w' = @{ Uri = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'; Id = 'OfficeXmlNamespaces.WordprocessingML'; Prefix = 'W' }
    'm' = @{ Uri = 'http://schemas.openxmlformats.org/officeDocument/2006/math'; Id = 'OfficeXmlNamespaces.Math'; Prefix = 'M' }
    'xml' = @{ Uri = 'http://www.w3.org/XML/1998/namespace'; Id = 'OfficeXmlNamespaces.Xml'; Prefix = 'Xml' }
    'mc' = @{ Uri = 'http://schemas.openxmlformats.org/markup-compatibility/2006'; Id = 'OfficeXmlNamespaces.MarkupCompatibility'; Prefix = 'Mc' }
    'wp' = @{ Uri = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'; Id = 'OfficeXmlNamespaces.WordprocessingDrawing'; Prefix = 'Wp' }
}
$xsdUri = 'http://www.w3.org/2001/XMLSchema'
$knownByUri = @{}
foreach ($entry in $namespaces.GetEnumerator()) { $knownByUri[$entry.Value.Uri] = $entry.Value }

# Index every global definition by '{namespace}name'.
$complexTypes = @{}
$simpleTypes = @{}
$groups = @{}
$attributeGroups = @{}
$elements = @{}
$attributes = @{}
foreach ($file in Get-ChildItem -LiteralPath $SchemaDirectory -Filter '*.xsd') {
    $document = [System.Xml.XmlDocument]::new()
    $document.Load($file.FullName)
    $schema = $document.DocumentElement
    $target = $schema.GetAttribute('targetNamespace')
    foreach ($node in $schema.ChildNodes) {
        if ($node.NamespaceURI -ne $xsdUri) { continue }
        $name = $node.GetAttribute('name')
        if (-not $name) { continue }
        $key = "{$target}$name"
        switch ($node.LocalName) {
            'complexType' { $complexTypes[$key] = $node }
            'simpleType' { $simpleTypes[$key] = $node }
            'group' { $groups[$key] = $node }
            'attributeGroup' { $attributeGroups[$key] = $node }
            'element' { $elements[$key] = $node }
            'attribute' { $attributes[$key] = $node }
        }
    }
}

function Resolve-QName([System.Xml.XmlElement] $Context, [string] $QName) {
    $colon = $QName.IndexOf(':')
    $prefix = if ($colon -lt 0) { '' } else { $QName.Substring(0, $colon) }
    $local = if ($colon -lt 0) { $QName } else { $QName.Substring($colon + 1) }
    $uri = $Context.GetNamespaceOfPrefix($prefix)
    if (-not $uri -and $prefix -eq '') { $uri = $Context.OwnerDocument.DocumentElement.GetAttribute('targetNamespace') }
    [pscustomobject]@{ Uri = $uri; Name = $local; Key = "{$uri}$local" }
}

function Get-TargetNamespace([System.Xml.XmlElement] $Node) { $Node.OwnerDocument.DocumentElement.GetAttribute('targetNamespace') }

# Maps a simple type to the value type a view exposes.
function Get-ValueType([System.Xml.XmlElement] $Context, [string] $QName) {
    if (-not $QName) { return 'string' }
    $type = Resolve-QName $Context $QName
    for ($depth = 0; $depth -lt 16; $depth++) {
        if ($type.Uri -eq $xsdUri) {
            switch ($type.Name) {
                'boolean' { return 'bool' }
                { $_ -in 'int', 'short', 'byte', 'integer' } { return 'int' }
                { $_ -in 'unsignedInt', 'unsignedShort', 'unsignedByte' } { return 'uint' }
                'long' { return 'long' }
                { $_ -in 'double', 'float', 'decimal' } { return 'double' }
                default { return 'string' }
            }
        }
        $definition = $simpleTypes[$type.Key]
        if (-not $definition) { return 'string' }
        $restriction = $definition.ChildNodes | Where-Object { $_.LocalName -eq 'restriction' } | Select-Object -First 1
        # Enumerations, unions and lists stay strings.
        if (-not $restriction -or ($restriction.ChildNodes | Where-Object { $_.LocalName -eq 'enumeration' })) { return 'string' }
        $type = Resolve-QName $restriction $restriction.GetAttribute('base')
    }
    'string'
}

# The child elements of a content model in document order, with the largest maxOccurs seen for each name. Each
# sequence item takes the next rank; every element under a choice shares the choice's rank, since the alternatives
# may appear in any order.
function Get-Children([System.Xml.XmlElement] $Node, [System.Collections.Generic.List[object]] $Into, [bool] $Repeated, [int] $Rank = -1) {
    foreach ($child in $Node.ChildNodes) {
        if ($child -isnot [System.Xml.XmlElement] -or $child.NamespaceURI -ne $xsdUri) { continue }
        $repeats = $Repeated -or ($child.GetAttribute('maxOccurs') -and $child.GetAttribute('maxOccurs') -ne '1')
        switch ($child.LocalName) {
            { $_ -in 'sequence', 'all' } { Get-Children $child $Into $repeats $Rank }
            'choice' {
                if ($script:orderedChoice -and $Rank -lt 0) {
                    Get-Children $child $Into $repeats $Rank
                } else {
                    $choiceRank = if ($Rank -ge 0) { $Rank } else { ($script:nextRank++) }
                    Get-Children $child $Into $repeats $choiceRank
                }
            }
            'group' {
                $groupKey = (Resolve-QName $child $child.GetAttribute('ref')).Key
                $group = $groups[$groupKey]
                if ($group) {
                    $outerOrdered = $script:orderedChoice
                    if ($orderedGroupKeys.Contains($groupKey)) { $script:orderedChoice = $true }
                    Get-Children $group $Into $repeats $Rank
                    $script:orderedChoice = $outerOrdered
                }
            }
            'element' {
                if ($child.GetAttribute('ref')) {
                    $reference = Resolve-QName $child $child.GetAttribute('ref')
                    $global = $elements[$reference.Key]
                    $entry = [pscustomobject]@{ Uri = $reference.Uri; Name = $reference.Name; Type = $null; Context = $global; Many = $repeats }
                    if ($global -and $global.GetAttribute('type')) { $entry.Type = Resolve-QName $global $global.GetAttribute('type') }
                } else {
                    $entry = [pscustomobject]@{ Uri = (Get-TargetNamespace $child); Name = $child.GetAttribute('name'); Type = $null; Context = $child; Many = $repeats }
                    if ($child.GetAttribute('type')) { $entry.Type = Resolve-QName $child $child.GetAttribute('type') }
                }
                $entry | Add-Member -NotePropertyName Rank -NotePropertyValue $(if ($Rank -ge 0) { $Rank } else { ($script:nextRank++) })
                $existing = $Into | Where-Object { $_.Uri -eq $entry.Uri -and $_.Name -eq $entry.Name } | Select-Object -First 1
                if ($existing) { $existing.Many = $true } else { $Into.Add($entry) }
            }
            { $_ -in 'complexContent', 'simpleContent' } { Get-Children $child $Into $Repeated $Rank }
            'extension' {
                $base = $complexTypes[(Resolve-QName $child $child.GetAttribute('base')).Key]
                if ($base) { Get-Children $base $Into $Repeated $Rank }
                Get-Children $child $Into $Repeated $Rank
            }
        }
    }
}

function Get-Attributes([System.Xml.XmlElement] $Node, [System.Collections.Generic.List[object]] $Into) {
    foreach ($child in $Node.ChildNodes) {
        if ($child -isnot [System.Xml.XmlElement] -or $child.NamespaceURI -ne $xsdUri) { continue }
        switch ($child.LocalName) {
            'attribute' {
                if ($child.GetAttribute('ref')) {
                    $reference = Resolve-QName $child $child.GetAttribute('ref')
                    $global = $attributes[$reference.Key]
                    $valueType = if ($global) { Get-ValueType $global $global.GetAttribute('type') } else { 'string' }
                    $Into.Add([pscustomobject]@{ Uri = $reference.Uri; Name = $reference.Name; ValueType = $valueType })
                } else {
                    $Into.Add([pscustomobject]@{ Uri = ''; Name = $child.GetAttribute('name'); ValueType = (Get-ValueType $child $child.GetAttribute('type')) })
                }
            }
            'attributeGroup' {
                $group = $attributeGroups[(Resolve-QName $child $child.GetAttribute('ref')).Key]
                if ($group) { Get-Attributes $group $Into }
            }
            { $_ -in 'complexContent', 'simpleContent' } { Get-Attributes $child $Into }
            'extension' {
                $base = $complexTypes[(Resolve-QName $child $child.GetAttribute('base')).Key]
                if ($base) { Get-Attributes $base $Into }
                Get-Attributes $child $Into
            }
            'restriction' { Get-Attributes $child $Into }
        }
    }
}

function ConvertTo-Pascal([string] $Name) { $Name.Substring(0, 1).ToUpperInvariant() + $Name.Substring(1) }

function Get-NamespaceExpression([string] $Uri) {
    if (-not $Uri) { return 'OfficeXmlNamespaces.None' }
    if ($knownByUri.ContainsKey($Uri)) { return $knownByUri[$Uri].Id }
    throw "No namespace id for '$Uri'."
}

$script:orderedChoice = $false
$orderedGroupKeys = [System.Collections.Generic.HashSet[string]]::new()
foreach ($group in $OrderedChoiceGroups) {
    $prefix, $name = $group.Split(':', 2)
    $info = $namespaces[$prefix]
    if (-not $info) { throw "Unknown prefix in '$group'." }
    $null = $orderedGroupKeys.Add("{$($info.Uri)}$name")
}

# View names for the requested types.
$viewNames = @{}
$requested = foreach ($type in $Types) {
    $prefix, $name = $type.Split(':', 2)
    $info = $namespaces[$prefix]
    if (-not $info) { throw "Unknown prefix in '$type'." }
    $key = "{$($info.Uri)}$name"
    if (-not $complexTypes.ContainsKey($key)) { throw "The schemas define no complex type '$type'." }
    $viewName = $info.Prefix + ($name -replace '^CT_', '') + 'View'
    $viewNames[$key] = $viewName
    [pscustomobject]@{ Key = $key; Name = $name; ViewName = $viewName; Node = $complexTypes[$key] }
}

$builder = [System.Text.StringBuilder]::new()
$null = $builder.AppendLine('// <auto-generated>')
$null = $builder.AppendLine("// Generated by $GeneratorName from the ISO/IEC 29500-4 transitional schemas.")
$null = $builder.AppendLine('// </auto-generated>')
$null = $builder.AppendLine('#nullable enable')
$null = $builder.AppendLine('using OfficeIMO.Core.Internal.OpenXml;')
if ($NodeNamespace) { $null = $builder.AppendLine("using $NodeNamespace;") }
$null = $builder.AppendLine()
$null = $builder.AppendLine("namespace $Namespace;")

foreach ($type in $requested) {
    $children = [System.Collections.Generic.List[object]]::new()
    $script:nextRank = 0
    Get-Children $type.Node $children $false
    $typeAttributes = [System.Collections.Generic.List[object]]::new()
    Get-Attributes $type.Node $typeAttributes

    $null = $builder.AppendLine()
    $null = $builder.AppendLine("/// <summary>View over an element of schema type <c>$($type.Name)</c>.</summary>")
    $null = $builder.AppendLine("internal readonly partial struct $($type.ViewName) {")
    if ($NodeType) {
        $null = $builder.AppendLine("    internal $($type.ViewName)($NodeType element) => Element = element;")
        $null = $builder.AppendLine()
        $null = $builder.AppendLine("    internal $($type.ViewName)(OfficeXmlTree tree, int node) : this(new $NodeType(tree, node)) { }")
        $null = $builder.AppendLine()
        $null = $builder.AppendLine("    internal $NodeType Element { get; }")
        $null = $builder.AppendLine('    internal OfficeXmlTree Tree => Element.Tree!;')
        $null = $builder.AppendLine('    internal int Node => Element.Node;')
    } else {
        $null = $builder.AppendLine("    internal $($type.ViewName)(OfficeXmlTree tree, int node) {")
        $null = $builder.AppendLine('        Tree = tree;')
        $null = $builder.AppendLine('        Node = node;')
        $null = $builder.AppendLine('    }')
        $null = $builder.AppendLine()
        $null = $builder.AppendLine('    internal OfficeXmlTree Tree { get; }')
        $null = $builder.AppendLine('    internal int Node { get; }')
    }

    $orderNamespaces = ($children | ForEach-Object { Get-NamespaceExpression $_.Uri }) -join ', '
    $orderNames = ($children | ForEach-Object { '"' + $_.Name + '"' }) -join ', '
    $orderRanks = ($children | ForEach-Object { $_.Rank }) -join ', '
    $null = $builder.AppendLine()
    $null = $builder.AppendLine('    /// <summary>Child element names in sequence order.</summary>')
    $null = $builder.AppendLine("    internal static readonly OfficeXmlChildOrder Order = new(new int[] { $orderNamespaces }, new string[] { $orderNames }, new int[] { $orderRanks });")

    $used = @{}
    foreach ($attribute in $typeAttributes) {
        $member = ConvertTo-Pascal $attribute.Name
        if ($attribute.Uri -eq $namespaces['r'].Uri) { $member = 'Relationship' + $member }
        $used[$member] = $true
        $ns = Get-NamespaceExpression $attribute.Uri
        $valueType = $attribute.ValueType
        $null = $builder.AppendLine()
        $null = $builder.AppendLine("    internal ${valueType}? $member {")
        $reader = switch ($valueType) { 'bool' { 'Boolean' } 'int' { 'Int32' } 'uint' { 'UInt32' } 'long' { 'Int64' } 'double' { 'Double' } default { '' } }
        if ($NodeType) {
            if ($valueType -eq 'string') {
                $null = $builder.AppendLine("        get => Element.Attribute($ns, `"$($attribute.Name)`");")
                $null = $builder.AppendLine("        set => Element.SetAttribute($ns, `"$($attribute.Name)`", value);")
            } else {
                $null = $builder.AppendLine("        get => OfficeXmlViewValues.Parse$reader(Element.Attribute($ns, `"$($attribute.Name)`"));")
                $null = $builder.AppendLine("        set => Element.SetAttribute($ns, `"$($attribute.Name)`", OfficeXmlViewValues.Format(value));")
            }
        } else {
            if ($valueType -eq 'string') {
                $null = $builder.AppendLine("        get => Tree.GetAttribute(Node, $ns, `"$($attribute.Name)`");")
            } else {
                $null = $builder.AppendLine("        get => OfficeXmlViewValues.Read$reader(Tree, Node, $ns, `"$($attribute.Name)`");")
            }
            $null = $builder.AppendLine("        set => OfficeXmlViewValues.Write(Tree, Node, $ns, `"$($attribute.Name)`", value);")
        }
        $null = $builder.AppendLine('    }')
    }

    foreach ($child in $children) {
        $member = ConvertTo-Pascal $child.Name
        if ($used.ContainsKey($member)) { $member += 'Element' }
        $used[$member] = $true
        $ns = Get-NamespaceExpression $child.Uri
        $childView = if ($child.Type) { $viewNames[$child.Type.Key] } else { $null }
        $isText = $child.Type -and ($child.Type.Uri -eq $xsdUri -or $simpleTypes.ContainsKey($child.Type.Key))
        if ($childView -and $NodeType) {
            $null = $builder.AppendLine()
            if ($child.Many) {
                $null = $builder.AppendLine("    internal $NodeType.ChildList ${member}List => Element.Children($ns, `"$($child.Name)`");")
                $null = $builder.AppendLine()
                $null = $builder.AppendLine("    /// <summary>Adds a <c>$($child.Name)</c> after the existing ones, in sequence order.</summary>")
                $null = $builder.AppendLine("    internal $childView Add$member() => new(Element.AddOrdered($ns, `"$($child.Name)`", Order));")
            } else {
                $null = $builder.AppendLine("    internal ${childView}? $member {")
                $null = $builder.AppendLine('        get {')
                $null = $builder.AppendLine("            $NodeType child = Element.Child($ns, `"$($child.Name)`");")
                $null = $builder.AppendLine("            return child.IsNone ? null : new $childView(child);")
                $null = $builder.AppendLine('        }')
                $null = $builder.AppendLine('    }')
                $null = $builder.AppendLine()
                $null = $builder.AppendLine("    internal $childView GetOrAdd$member() => new(Element.GetOrAddChild($ns, `"$($child.Name)`", Order));")
            }
        } elseif ($isText -and -not $child.Many -and $NodeType) {
            $null = $builder.AppendLine()
            $null = $builder.AppendLine("    internal string? $member {")
            $null = $builder.AppendLine("        get => Element.ChildText($ns, `"$($child.Name)`");")
            $null = $builder.AppendLine("        set => Element.SetChildText($ns, `"$($child.Name)`", value, Order);")
            $null = $builder.AppendLine('    }')
        } elseif ($childView) {
            $null = $builder.AppendLine()
            if ($child.Many) {
                $null = $builder.AppendLine("    internal OfficeXmlElementList ${member}List => new(Tree, Node, $ns, `"$($child.Name)`");")
                $null = $builder.AppendLine()
                $null = $builder.AppendLine("    /// <summary>Adds a <c>$($child.Name)</c> after the existing ones, in sequence order.</summary>")
                $null = $builder.AppendLine("    internal $childView Add$member() {")
                $null = $builder.AppendLine("        int child = Tree.CreateElement($ns, `"$($child.Name)`");")
                $null = $builder.AppendLine('        Tree.InsertOrdered(Node, child, Order);')
                $null = $builder.AppendLine("        return new $childView(Tree, child);")
                $null = $builder.AppendLine('    }')
            } else {
                $null = $builder.AppendLine("    internal ${childView}? $member {")
                $null = $builder.AppendLine('        get {')
                $null = $builder.AppendLine("            int child = Tree.FindChild(Node, $ns, `"$($child.Name)`");")
                $null = $builder.AppendLine("            return child == OfficeXmlTree.NoNode ? null : new $childView(Tree, child);")
                $null = $builder.AppendLine('        }')
                $null = $builder.AppendLine('    }')
                $null = $builder.AppendLine()
                $null = $builder.AppendLine("    internal $childView GetOrAdd$member() => new(Tree, Tree.GetOrInsertChild(Node, $ns, `"$($child.Name)`", Order));")
            }
        } elseif ($isText -and -not $child.Many) {
            $null = $builder.AppendLine()
            $null = $builder.AppendLine("    internal string? $member {")
            $null = $builder.AppendLine("        get => OfficeXmlViewValues.ReadChildText(Tree, Node, $ns, `"$($child.Name)`");")
            $null = $builder.AppendLine("        set => OfficeXmlViewValues.WriteChildText(Tree, Node, $ns, `"$($child.Name)`", value, Order);")
            $null = $builder.AppendLine('    }')
        }
    }
    $null = $builder.AppendLine('}')
}

$text = $builder.ToString().Replace("`r`n", "`n")
[System.IO.File]::WriteAllText($OutputPath, $text, [System.Text.UTF8Encoding]::new($false))
