namespace OfficeIMO.Workflows;

/// <summary>Supported ONIX list 158 resource types, excluding transitional and deprecated codes.</summary>
public enum BookOnixResourceContentType {
    /// <summary>Front Cover (01).</summary>
    FrontCover = 1,
    /// <summary>Back Cover (02).</summary>
    BackCover = 2,
    /// <summary>Cover Or Pack (03).</summary>
    CoverOrPack = 3,
    /// <summary>Contributor Picture (04).</summary>
    ContributorPicture = 4,
    /// <summary>Collection Artwork (05).</summary>
    CollectionArtwork = 5,
    /// <summary>Collection Logo (06).</summary>
    CollectionLogo = 6,
    /// <summary>Product Artwork (07).</summary>
    ProductArtwork = 7,
    /// <summary>Product Logo (08).</summary>
    ProductLogo = 8,
    /// <summary>Publisher Logo (09).</summary>
    PublisherLogo = 9,
    /// <summary>Imprint Logo (10).</summary>
    ImprintLogo = 10,
    /// <summary>Contributor Interview (11).</summary>
    ContributorInterview = 11,
    /// <summary>Contributor Presentation (12).</summary>
    ContributorPresentation = 12,
    /// <summary>Contributor Reading (13).</summary>
    ContributorReading = 13,
    /// <summary>Contributor Event Schedule (14).</summary>
    ContributorEventSchedule = 14,
    /// <summary>Sample Content (15).</summary>
    SampleContent = 15,
    /// <summary>Widget (16).</summary>
    Widget = 16,
    /// <summary>Review (17).</summary>
    Review = 17,
    /// <summary>Commentary (18).</summary>
    Commentary = 18,
    /// <summary>Reading Group Guide (19).</summary>
    ReadingGroupGuide = 19,
    /// <summary>Teacher Guide (20).</summary>
    TeacherGuide = 20,
    /// <summary>Feature Article (21).</summary>
    FeatureArticle = 21,
    /// <summary>Character Interview (22).</summary>
    CharacterInterview = 22,
    /// <summary>Wallpaper (23).</summary>
    Wallpaper = 23,
    /// <summary>Press Release (24).</summary>
    PressRelease = 24,
    /// <summary>Table Of Contents (25).</summary>
    TableOfContents = 25,
    /// <summary>Trailer (26).</summary>
    Trailer = 26,
    /// <summary>Full Content (28).</summary>
    FullContent = 28,
    /// <summary>Full Cover (29).</summary>
    FullCover = 29,
    /// <summary>Master Brand Logo (30).</summary>
    MasterBrandLogo = 30,
    /// <summary>Description (31).</summary>
    Description = 31,
    /// <summary>Index (32).</summary>
    Index = 32,
    /// <summary>Student Guide (33).</summary>
    StudentGuide = 33,
    /// <summary>Publisher Catalog (34).</summary>
    PublisherCatalog = 34,
    /// <summary>Advertisement Panel (35).</summary>
    AdvertisementPanel = 35,
    /// <summary>Advertisement Page (36).</summary>
    AdvertisementPage = 36,
    /// <summary>Promotional Event Material (37).</summary>
    PromotionalEventMaterial = 37,
    /// <summary>Digital Review Copy (38).</summary>
    DigitalReviewCopy = 38,
    /// <summary>Instructional Material (39).</summary>
    InstructionalMaterial = 39,
    /// <summary>Errata (40).</summary>
    Errata = 40,
    /// <summary>Introduction (41).</summary>
    Introduction = 41,
    /// <summary>Collection Description (42).</summary>
    CollectionDescription = 42,
    /// <summary>Bibliography (43).</summary>
    Bibliography = 43,
    /// <summary>Abstract (44).</summary>
    Abstract = 44,
    /// <summary>Cover Holding Image (45).</summary>
    CoverHoldingImage = 45,
    /// <summary>Rules Or Instructions (46).</summary>
    RulesOrInstructions = 46,
    /// <summary>Transcript (47).</summary>
    Transcript = 47,
    /// <summary>Cast And Credits (48).</summary>
    CastAndCredits = 48,
    /// <summary>Social Media Image (49).</summary>
    SocialMediaImage = 49,
    /// <summary>Supplementary Learning Resources (50).</summary>
    SupplementaryLearningResources = 50,
    /// <summary>Cover Flap Image (51).</summary>
    CoverFlapImage = 51,
    /// <summary>Warning Label (52).</summary>
    WarningLabel = 52,
    /// <summary>Page Edge Image (54).</summary>
    PageEdgeImage = 54,
    /// <summary>Endpaper Image (55).</summary>
    EndpaperImage = 55,
    /// <summary>Spine Image (56).</summary>
    SpineImage = 56,
    /// <summary>Spine Panorama Image (57).</summary>
    SpinePanoramaImage = 57,
}

/// <summary>ONIX list 159 resource modes.</summary>
public enum BookOnixResourceMode {
    /// <summary>Application (01).</summary>
    Application = 1,
    /// <summary>Audio (02).</summary>
    Audio = 2,
    /// <summary>Image (03).</summary>
    Image = 3,
    /// <summary>Text (04).</summary>
    Text = 4,
    /// <summary>Video (05).</summary>
    Video = 5,
    /// <summary>Multi Mode (06).</summary>
    MultiMode = 6,
}

/// <summary>ONIX list 161 hosting and delivery forms. These are assertions, not executed operations.</summary>
public enum BookOnixResourceForm {
    /// <summary>The sender hosts the resource for linking (01).</summary>
    Linkable = 1,
    /// <summary>The recipient downloads and hosts a copy (02).</summary>
    Downloadable = 2,
    /// <summary>An application supplied for embedding (03).</summary>
    EmbeddableApplication = 3
}
