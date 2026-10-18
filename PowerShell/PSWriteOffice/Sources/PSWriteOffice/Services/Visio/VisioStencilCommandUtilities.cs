using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Visio.Stencils;

namespace PSWriteOffice.Services.Visio;

internal static class VisioStencilCommandUtilities
{
    private static readonly string[] PackageExtensions =
    {
        ".vsdx",
        ".vssx",
        ".vstx",
        ".vsdm",
        ".vssm",
        ".vstm"
    };

    internal static VisioStencilCatalog GetBuiltInCatalog(OfficeVisioBuiltInStencilCatalog builtIn)
    {
        return builtIn switch
        {
            OfficeVisioBuiltInStencilCatalog.BasicShapes => VisioStencils.BasicShapes,
            OfficeVisioBuiltInStencilCatalog.Flowchart => VisioStencils.Flowchart,
            OfficeVisioBuiltInStencilCatalog.BlockDiagram => VisioStencils.BlockDiagram,
            OfficeVisioBuiltInStencilCatalog.Architecture => VisioStencils.Architecture,
            OfficeVisioBuiltInStencilCatalog.Network => VisioStencils.Network,
            OfficeVisioBuiltInStencilCatalog.Infrastructure => VisioStencils.Infrastructure,
            OfficeVisioBuiltInStencilCatalog.Cloud => VisioStencils.Cloud,
            OfficeVisioBuiltInStencilCatalog.SecurityIdentity => VisioStencils.SecurityIdentity,
            OfficeVisioBuiltInStencilCatalog.ContainersKubernetes => VisioStencils.ContainersKubernetes,
            OfficeVisioBuiltInStencilCatalog.DataPlatform => VisioStencils.DataPlatform,
            OfficeVisioBuiltInStencilCatalog.CollaborationBusiness => VisioStencils.CollaborationBusiness,
            OfficeVisioBuiltInStencilCatalog.Swimlane => VisioStencils.Swimlane,
            OfficeVisioBuiltInStencilCatalog.OrgChart => VisioStencils.OrgChart,
            OfficeVisioBuiltInStencilCatalog.Timeline => VisioStencils.Timeline,
            OfficeVisioBuiltInStencilCatalog.Sequence => VisioStencils.Sequence,
            _ => VisioStencils.All
        };
    }

    internal static VisioStencilCatalog LoadCatalog(
        PSCmdlet cmdlet,
        IEnumerable<string> paths,
        VisioStencilPackageLoadOptions options,
        bool recursive)
    {
        if (paths == null)
        {
            throw new ArgumentNullException(nameof(paths));
        }

        if (options == null)
        {
            throw new ArgumentNullException(nameof(options));
        }

        var packagePaths = new List<string>();
        var manifestPaths = new List<string>();

        foreach (var inputPath in paths.Where(path => !string.IsNullOrWhiteSpace(path)))
        {
            var resolvedPath = VisioCommandUtilities.ResolvePath(cmdlet, inputPath);
            if (Directory.Exists(resolvedPath))
            {
                packagePaths.AddRange(VisioStencilPackageCatalog.EnumeratePackageFiles(resolvedPath, recursive));
                continue;
            }

            if (!File.Exists(resolvedPath))
            {
                throw new FileNotFoundException("Stencil catalog path was not found.", resolvedPath);
            }

            if (IsPackagePath(resolvedPath))
            {
                packagePaths.Add(resolvedPath);
            }
            else
            {
                manifestPaths.Add(resolvedPath);
            }
        }

        if (manifestPaths.Count > 0 && packagePaths.Count > 0)
        {
            throw new PSArgumentException("Do not mix Visio packages and OfficeIMO native stencil manifest files in one call.", nameof(paths));
        }

        if (manifestPaths.Count > 1)
        {
            throw new PSArgumentException("Only one OfficeIMO native stencil manifest can be loaded at a time.", nameof(paths));
        }

        if (manifestPaths.Count == 1)
        {
            return VisioStencilCatalog.Load(manifestPaths[0]);
        }

        return VisioStencilPackageCatalog.LoadMany(packagePaths, options);
    }

    internal static VisioStencilPackageLoadOptions BuildPackageLoadOptions(
        string? catalogName,
        string? category,
        string? idPrefix,
        string[]? masterName,
        bool includeUnsupportedMasters,
        bool noLearnMasterDimensions,
        bool noPreviewImageMetadata,
        bool noConnectionPointMetadata,
        double defaultWidth,
        double defaultHeight)
    {
        return new VisioStencilPackageLoadOptions
        {
            CatalogName = catalogName,
            Category = category,
            IdPrefix = idPrefix,
            MasterNames = masterName,
            IncludeUnsupportedMasters = includeUnsupportedMasters,
            LearnMasterDimensions = !noLearnMasterDimensions,
            ExtractPreviewImageMetadata = !noPreviewImageMetadata,
            ExtractConnectionPointMetadata = !noConnectionPointMetadata,
            DefaultWidth = defaultWidth,
            DefaultHeight = defaultHeight
        };
    }

    private static bool IsPackagePath(string path)
    {
        var extension = Path.GetExtension(path);
        return PackageExtensions.Any(candidate => string.Equals(candidate, extension, StringComparison.OrdinalIgnoreCase));
    }
}
