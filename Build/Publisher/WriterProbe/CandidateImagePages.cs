using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Projects two distinct public-model assets on separated physical pages.</summary>
internal static class CandidateImagePages {
    internal static PublisherNativeCreationPlan Create(byte[] firstImage, byte[] lastImage, bool includeNativeNames = false,
        OfficeIMO.Drawing.OfficeImageSourceCrop crop = default, bool shareFirstImage = false) =>
        PublisherNativeCreationProjection.SingleMasterContent(Author(firstImage, lastImage, crop, shareFirstImage), includeNativeNames);

    internal static PublisherDocument Author(byte[] firstImage, byte[] lastImage,
        OfficeIMO.Drawing.OfficeImageSourceCrop crop = default, bool shareFirstImage = false) {
        PublisherDocument document = PublisherDocument.Create(612, 792);
        PublisherPage master = document.AddMasterPage("Master Page A");
        document.SetPageMaster(document.Pages[0].Id, master.Id);
        PublisherPicture first = document.AddPicture(document.Pages[0].Id, firstImage, 72, 144, 96, 96);
        if (crop.HasCrop) document.SetPictureCrop(first.Id, crop);
        document.AddPage(master.Id);
        PublisherPage lastPage = document.AddPage(master.Id);
        if (shareFirstImage) document.AddPicture(lastPage.Id, first.ImageId!.Value, 72, 144, 96, 96);
        document.AddPicture(lastPage.Id, lastImage, 216, 144, 192, 96);
        return document;
    }
}
