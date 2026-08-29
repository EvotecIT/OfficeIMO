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
    Article
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
    Other
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
    Other
}
