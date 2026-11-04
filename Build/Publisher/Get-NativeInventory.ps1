function Get-NativeInventory {
    <#
    .SYNOPSIS
    Captures physical/master page assignments and current native object observations.
    .PARAMETER Publication
    The native publication opened by the isolated comparison worker.
    .PARAMETER Objects
    Retains observed COM wrappers for the isolated worker's lifetime.
    .OUTPUTS
    Publication inventory with separate physical and master pages and read-only state.
    #>
    param($Publication, [Collections.Generic.List[object]] $Objects)
    $nativePages = [Collections.Generic.List[object]]::new()
    for ($pageIndex = 1; $pageIndex -le $Publication.Pages.Count; $pageIndex++) {
        $currentPage = $Publication.Pages.Item($pageIndex)
        $objects.Add($currentPage)
        $masterInfo = $null
        $masterStatus = 'None'
        $masterError = $null
        try {
            $master = $currentPage.Master
            if ($null -ne $master) {
                $objects.Add($master)
                $masterInfo = [pscustomobject]@{ pageId = [int]$master.PageID; name = [string]$master.Name }
                $masterStatus = 'Observed'
            }
        } catch {
            $masterStatus = 'Unavailable'
            $masterError = $_.Exception.HResult
        }
        $nativePages.Add([pscustomobject]@{
            index = $pageIndex
            pageId = [int]$currentPage.PageID
            pageIndex = [int]$currentPage.PageIndex
            pageNumber = [string]$currentPage.PageNumber
            pageType = [int]$currentPage.PageType
            name = [string]$currentPage.Name
            masterStatus = $masterStatus
            master = $masterInfo
            masterErrorHResult = $masterError
            ignoreMaster = [bool]$currentPage.IgnoreMaster
            shapes = @(Get-NativeShapeInventory -Page $currentPage -Objects $Objects)
        })
    }
    $nativeMasters = [Collections.Generic.List[object]]::new()
    for ($masterIndex = 1; $masterIndex -le $Publication.MasterPages.Count; $masterIndex++) {
        $masterPage = $Publication.MasterPages.Item($masterIndex)
        $objects.Add($masterPage)
        # Page.Master is intentionally not read on a master: the documented API
        # rejects it. Native master definitions remain separate from Pages.
        $nativeMasters.Add([pscustomobject]@{
            index = $masterIndex
            pageId = [int]$masterPage.PageID
            pageIndex = [int]$masterPage.PageIndex
            pageType = [int]$masterPage.PageType
            name = [string]$masterPage.Name
            isTwoPageMaster = [bool]$masterPage.IsTwoPageMaster
            shapeCount = [int]$masterPage.Shapes.Count
            shapes = @(Get-NativeShapeInventory -Page $masterPage -Objects $Objects)
        })
    }
    return [pscustomobject]@{
        pageCount = [int]$Publication.Pages.Count
        widthPoints = [double]$Publication.PageSetup.PageWidth
        heightPoints = [double]$Publication.PageSetup.PageHeight
        readOnly = [bool]$Publication.ReadOnly
        defaultTabStopPoints = [double]$Publication.DefaultTabStop
        pages = $nativePages.ToArray()
        masterPages = $nativeMasters.ToArray()
    }
}
