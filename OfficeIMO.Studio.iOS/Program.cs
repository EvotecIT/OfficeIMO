using UIKit;
using System.Diagnostics.CodeAnalysis;

namespace OfficeIMO.Studio.iOS;

internal static class Program {
    [DynamicDependency(DynamicallyAccessedMemberTypes.All, typeof(AppDelegate))]
    private static void Main(string[] args) => UIApplication.Main(args, null, typeof(AppDelegate));
}
