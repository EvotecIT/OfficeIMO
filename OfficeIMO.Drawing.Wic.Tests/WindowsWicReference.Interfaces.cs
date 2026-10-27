using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;
using OfficeIMO.Drawing.Wic.Interop;

namespace OfficeIMO.Drawing.Wic.Tests;

[GeneratedComInterface(Options = ComInterfaceOptions.ComObjectWrapper)]
[Guid("3b16811b-6a43-4ec9-a813-3d930c13b940")]
internal partial interface IWicReferenceFrame : IWicBitmapSource {
    [PreserveSig] int GetMetadataQueryReader(out nint reader);
    [PreserveSig] int GetColorContexts(uint count, nint contexts, out uint actual);
    [PreserveSig] int GetThumbnail(out nint thumbnail);
}

[GeneratedComInterface(Options = ComInterfaceOptions.ComObjectWrapper, StringMarshalling = StringMarshalling.Utf16)]
[Guid("30989668-e1c9-4597-b395-458eedb808df")]
internal partial interface IWicReferenceMetadata {
    [PreserveSig] int GetContainerFormat(out Guid format);
    [PreserveSig] int GetLocation(uint length, nint buffer, out uint actual);
    [PreserveSig] int GetMetadataByName(string name, ref WicPropVariant value);
}

[GeneratedComInterface(Options = ComInterfaceOptions.ComObjectWrapper)]
[Guid("3c613a02-34b2-44ea-9a7c-45aea9c6fd6d")]
internal partial interface IWicReferenceColorContext {
    [PreserveSig] int InitializeFromFilename(nint filename);
    [PreserveSig] int InitializeFromMemory(nint bytes, uint length);
    [PreserveSig] int InitializeFromExifColorSpace(uint value);
}

[GeneratedComInterface(Options = ComInterfaceOptions.ComObjectWrapper)]
[Guid("b66f034f-d0e2-40ab-b436-6de39e321a94")]
internal partial interface IWicReferenceColorTransform : IWicBitmapSource {
    [PreserveSig] int Initialize(nint source, nint sourceContext, nint destinationContext, in Guid format);
}

// PROPVARIANT is 24 bytes on 64-bit Windows, 16 on 32-bit. Overallocate on x86;
// the fixed scalar fields used here retain their documented offsets on both.
[StructLayout(LayoutKind.Explicit, Size = 24)]
internal struct WicPropVariant {
    [FieldOffset(0)] internal ushort Type;
    [FieldOffset(8)] internal ushort UInt16;
}
