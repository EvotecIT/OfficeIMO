using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Projects two distinct public-model assets on separated physical pages.</summary>
internal static class CandidateImagePages {
    internal static PublisherNativeCreationPlan Create(byte[] firstImage, byte[] lastImage, bool includeNativeNames = false,
        OfficeIMO.Drawing.OfficeImageSourceCrop crop = default, bool shareFirstImage = false) =>
        PublisherNativeCreationProjection.SingleMasterContent(Author(firstImage, lastImage, crop, shareFirstImage), includeNativeNames);

    internal static PublisherDocument Author(byte[] firstImage, byte[] lastImage,
        OfficeIMO.Drawing.OfficeImageSourceCrop crop = default, bool shareFirstImage = false) {
        PublisherDocument document = PublisherDocument.Create(OfficeLength.Points(612), OfficeLength.Points(792));
        PublisherPage master = document.AddMasterPage("Master Page A");
        document.SetPageMaster(document.Pages[0].Id, master.Id);
        PublisherPicture first = document.AddImage(document.Pages[0].Id, firstImage, OfficeLength.Points(72), OfficeLength.Points(144), OfficeLength.Points(96), OfficeLength.Points(96));
        if (crop.HasCrop) document.SetPictureCrop(first.Id, crop);
        document.AddPage(master.Id);
        PublisherPage lastPage = document.AddPage(master.Id);
        if (shareFirstImage) document.AddImage(lastPage.Id, first.ImageId!.Value, OfficeLength.Points(72), OfficeLength.Points(144), OfficeLength.Points(96), OfficeLength.Points(96));
        document.AddImage(lastPage.Id, lastImage, OfficeLength.Points(216), OfficeLength.Points(144), OfficeLength.Points(192), OfficeLength.Points(96));
        return document;
    }
}
