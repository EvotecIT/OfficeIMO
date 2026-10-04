namespace OfficeIMO.Bibliography;

/// <summary>Supported bibliography interchange formats.</summary>
public enum BibliographyFormat {
    /// <summary>Classic BibTeX database syntax.</summary>
    BibTex = 0,
    /// <summary>BibLaTeX database syntax.</summary>
    BibLatex,
    /// <summary>Citation Style Language JSON data.</summary>
    CslJson,
    /// <summary>Research Information Systems tagged data.</summary>
    Ris,
    /// <summary>PubMed NBIB/MEDLINE tagged data.</summary>
    Nbib,
    /// <summary>EndNote XML interchange data.</summary>
    EndNoteXml
}

/// <summary>Format-neutral bibliography item kinds.</summary>
public enum BibliographyItemType {
    /// <summary>Type was not recognized.</summary>
    Unknown = 0,
    /// <summary>Journal article.</summary>
    ArticleJournal,
    /// <summary>Magazine article.</summary>
    ArticleMagazine,
    /// <summary>Newspaper article.</summary>
    ArticleNewspaper,
    /// <summary>Book.</summary>
    Book,
    /// <summary>Chapter or contribution in a book.</summary>
    Chapter,
    /// <summary>Conference paper.</summary>
    PaperConference,
    /// <summary>Conference proceedings.</summary>
    Proceedings,
    /// <summary>Report.</summary>
    Report,
    /// <summary>Thesis or dissertation.</summary>
    Thesis,
    /// <summary>Web page.</summary>
    WebPage,
    /// <summary>Dataset.</summary>
    Dataset,
    /// <summary>Software.</summary>
    Software,
    /// <summary>Patent.</summary>
    Patent,
    /// <summary>Legal case.</summary>
    LegalCase,
    /// <summary>Manuscript or other unpublished work.</summary>
    Manuscript,
    /// <summary>Personal communication.</summary>
    PersonalCommunication,
    /// <summary>Generic document.</summary>
    Document,
    /// <summary>Generic article without a more specific journal, magazine, or newspaper classification.</summary>
    Article,
    /// <summary>Legislative bill.</summary>
    Bill,
    /// <summary>Radio or television broadcast.</summary>
    Broadcast,
    /// <summary>Classic work.</summary>
    Classic,
    /// <summary>Collection of works or records.</summary>
    Collection,
    /// <summary>Generic reference-work entry.</summary>
    Entry,
    /// <summary>Dictionary entry.</summary>
    EntryDictionary,
    /// <summary>Encyclopedia entry.</summary>
    EntryEncyclopedia,
    /// <summary>Event.</summary>
    Event,
    /// <summary>Figure.</summary>
    Figure,
    /// <summary>Graphic work.</summary>
    Graphic,
    /// <summary>Hearing.</summary>
    Hearing,
    /// <summary>Interview.</summary>
    Interview,
    /// <summary>Legislation.</summary>
    Legislation,
    /// <summary>Map.</summary>
    Map,
    /// <summary>Motion picture.</summary>
    MotionPicture,
    /// <summary>Musical score.</summary>
    MusicalScore,
    /// <summary>Pamphlet.</summary>
    Pamphlet,
    /// <summary>Performance.</summary>
    Performance,
    /// <summary>Periodical as a whole.</summary>
    Periodical,
    /// <summary>Generic post.</summary>
    Post,
    /// <summary>Weblog post.</summary>
    PostWeblog,
    /// <summary>Regulation.</summary>
    Regulation,
    /// <summary>Review.</summary>
    Review,
    /// <summary>Book review.</summary>
    ReviewBook,
    /// <summary>Song.</summary>
    Song,
    /// <summary>Speech.</summary>
    Speech,
    /// <summary>Standard.</summary>
    Standard,
    /// <summary>Treaty.</summary>
    Treaty
}

/// <summary>Contributor roles shared by supported formats.</summary>
public enum BibliographyContributorRole {
    /// <summary>Author.</summary>
    Author = 0,
    /// <summary>Editor.</summary>
    Editor,
    /// <summary>Translator.</summary>
    Translator,
    /// <summary>Recipient.</summary>
    Recipient,
    /// <summary>Interviewer.</summary>
    Interviewer,
    /// <summary>Composer.</summary>
    Composer,
    /// <summary>Collection editor.</summary>
    CollectionEditor,
    /// <summary>Other contributor role.</summary>
    Other,
    /// <summary>Chair.</summary>
    Chair,
    /// <summary>Compiler.</summary>
    Compiler,
    /// <summary>Author of the containing work.</summary>
    ContainerAuthor,
    /// <summary>Contributor without a more specific role.</summary>
    Contributor,
    /// <summary>Curator.</summary>
    Curator,
    /// <summary>Director.</summary>
    Director,
    /// <summary>Editorial director.</summary>
    EditorialDirector,
    /// <summary>Executive producer.</summary>
    ExecutiveProducer,
    /// <summary>Guest.</summary>
    Guest,
    /// <summary>Host.</summary>
    Host,
    /// <summary>Illustrator.</summary>
    Illustrator,
    /// <summary>Narrator.</summary>
    Narrator,
    /// <summary>Organizer.</summary>
    Organizer,
    /// <summary>Original author.</summary>
    OriginalAuthor,
    /// <summary>Performer.</summary>
    Performer,
    /// <summary>Producer.</summary>
    Producer,
    /// <summary>Author of the reviewed work.</summary>
    ReviewedAuthor,
    /// <summary>Script writer.</summary>
    ScriptWriter,
    /// <summary>Series creator.</summary>
    SeriesCreator
}

/// <summary>Date roles shared by supported formats.</summary>
public enum BibliographyDateRole {
    /// <summary>Issued or published date.</summary>
    Issued = 0,
    /// <summary>Accessed date.</summary>
    Accessed,
    /// <summary>Submitted date.</summary>
    Submitted,
    /// <summary>Original publication date.</summary>
    Original,
    /// <summary>Event date.</summary>
    Event,
    /// <summary>Other date.</summary>
    Other,
    /// <summary>Date the work became available.</summary>
    Available
}
