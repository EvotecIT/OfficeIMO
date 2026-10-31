namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Projects two distinct public-model assets on separated physical pages.</summary>
internal static class CandidateImagePages {
    internal static CandidatePublicationSpec Create(byte[] firstImage, byte[] lastImage, bool includeNativeNames = false,
        OfficeIMO.Drawing.OfficeImageSourceCrop crop = default) {
        PublisherDocument document = PublisherDocument.Create(612, 792);
        PublisherPage master = document.AddMasterPage("Master Page A");
        document.SetPageMaster(document.Pages[0].Id, master.Id);
        PublisherPicture first = document.AddPicture(document.Pages[0].Id, firstImage, 72, 144, 96, 96);
        if (crop.HasCrop) {
            document.SetPictureCrop(first.Id, crop);
            first = document.Pages[0].Pictures.Single();
        }
        document.AddPage(master.Id);
        PublisherPage lastPage = document.AddPage(master.Id);
        PublisherPicture last = document.AddPicture(lastPage.Id, lastImage, 216, 144, 192, 96);
        CandidateObjectSpec Picture(PublisherPicture picture, int pageIndex) => new(CandidateObjectKind.Picture,
            new CandidatePlacement(picture.X, picture.Y, picture.Width, picture.Height),
            Image: document.Images.Single(asset => asset.Id == picture.ImageId).GetBytes(),
            PageIndex: pageIndex, Paint: picture.CreatePaint(),
            ImageName: includeNativeNames ? (pageIndex == 0 ? "Image1.png" : "Image2.png") : null, Crop: picture.Crop);
        return new CandidatePublicationSpec(checked((uint)(document.Pages[0].Width * 12700)),
            checked((uint)(document.Pages[0].Height * 12700)), new[] { Picture(first, 0), Picture(last, 2) }, document.Pages.Count);
    }
}
