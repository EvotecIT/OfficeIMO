namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Projects two distinct public-model assets on separated physical pages.</summary>
internal static class CandidateImagePages {
    internal static CandidatePublicationSpec Create(byte[] firstImage, byte[] lastImage) {
        PublisherDocument document = PublisherDocument.Create(612, 792);
        PublisherPage master = document.AddMasterPage("Master Page A");
        document.SetPageMaster(document.Pages[0].Id, master.Id);
        PublisherPicture first = document.AddPicture(document.Pages[0].Id, firstImage, 72, 144, 96, 96);
        document.AddPage(master.Id);
        PublisherPage lastPage = document.AddPage(master.Id);
        PublisherPicture last = document.AddPicture(lastPage.Id, lastImage, 216, 144, 192, 96);
        CandidateObjectSpec Picture(PublisherPicture picture, int pageIndex) => new(CandidateObjectKind.Picture,
            new CandidatePlacement(picture.X, picture.Y, picture.Width, picture.Height),
            Image: document.Images.Single(asset => asset.Id == picture.ImageId).GetBytes(),
            PageIndex: pageIndex, Paint: picture.CreatePaint());
        return new CandidatePublicationSpec(checked((uint)(document.Pages[0].Width * 12700)),
            checked((uint)(document.Pages[0].Height * 12700)), new[] { Picture(first, 0), Picture(last, 2) }, document.Pages.Count);
    }
}
