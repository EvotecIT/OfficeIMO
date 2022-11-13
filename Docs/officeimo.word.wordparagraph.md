# WordParagraph

Namespace: OfficeIMO.Word

```csharp
public class WordParagraph
```

Inheritance [Object](https://docs.microsoft.com/en-us/dotnet/api/system.object) → [WordParagraph](./officeimo.word.wordparagraph.md)

## Properties

### **IsLastRun**

```csharp
public bool IsLastRun { get; }
```

#### Property Value

[Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

### **IsFirstRun**

```csharp
public bool IsFirstRun { get; }
```

#### Property Value

[Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

### **Image**

```csharp
public WordImage Image { get; }
```

#### Property Value

[WordImage](./officeimo.word.wordimage.md)<br>

### **IsListItem**

```csharp
public bool IsListItem { get; }
```

#### Property Value

[Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

### **ListItemLevel**

```csharp
public Nullable<int> ListItemLevel { get; set; }
```

#### Property Value

[Nullable&lt;Int32&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **Style**

```csharp
public Nullable<WordParagraphStyles> Style { get; set; }
```

#### Property Value

[Nullable&lt;WordParagraphStyles&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **Text**

Get or set a text within Paragraph

```csharp
public string Text { get; set; }
```

#### Property Value

[String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

### **PageBreak**

Get PageBreaks within Paragraph

```csharp
public WordBreak PageBreak { get; }
```

#### Property Value

[WordBreak](./officeimo.word.wordbreak.md)<br>

### **Break**

Get Breaks within Paragraph

```csharp
public WordBreak Break { get; }
```

#### Property Value

[WordBreak](./officeimo.word.wordbreak.md)<br>

### **Bookmark**

```csharp
public WordBookmark Bookmark { get; }
```

#### Property Value

[WordBookmark](./officeimo.word.wordbookmark.md)<br>

### **Equation**

```csharp
public WordEquation Equation { get; }
```

#### Property Value

[WordEquation](./officeimo.word.wordequation.md)<br>

### **Field**

```csharp
public WordField Field { get; }
```

#### Property Value

[WordField](./officeimo.word.wordfield.md)<br>

### **Hyperlink**

```csharp
public WordHyperLink Hyperlink { get; }
```

#### Property Value

[WordHyperLink](./officeimo.word.wordhyperlink.md)<br>

### **IsHyperLink**

```csharp
public bool IsHyperLink { get; }
```

#### Property Value

[Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

### **IsField**

```csharp
public bool IsField { get; }
```

#### Property Value

[Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

### **IsBookmark**

```csharp
public bool IsBookmark { get; }
```

#### Property Value

[Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

### **IsEquation**

```csharp
public bool IsEquation { get; }
```

#### Property Value

[Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

### **IsStructuredDocumentTag**

```csharp
public bool IsStructuredDocumentTag { get; }
```

#### Property Value

[Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

### **IsImage**

```csharp
public bool IsImage { get; }
```

#### Property Value

[Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

### **ParagraphAlignment**

Alignment aka Paragraph Alignment. This element specifies the paragraph alignment which shall be applied to text in this paragraph.
 If this element is omitted on a given paragraph, its value is determined by the setting previously set at any level of the style hierarchy (i.e.that previous setting remains unchanged). If this setting is never specified in the style hierarchy, then no alignment is applied to the paragraph.

```csharp
public Nullable<JustificationValues> ParagraphAlignment { get; set; }
```

#### Property Value

[Nullable&lt;JustificationValues&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **VerticalCharacterAlignmentOnLine**

Text Alignment aka Vertical Character Alignment on Line. This element specifies the vertical alignment of all text on each line displayed within a paragraph. If the line height (before any added spacing) is larger than one or more characters on the line, all characters are aligned to each other as specified by this element.
 If this element is omitted on a given paragraph, its value is determined by the setting previously set at any level of the style hierarchy (i.e.that previous setting remains unchanged). If this setting is never specified in the style hierarchy, then the vertical alignment of all characters on the line shall be automatically determined by the consumer.

```csharp
public Nullable<VerticalTextAlignmentValues> VerticalCharacterAlignmentOnLine { get; set; }
```

#### Property Value

[Nullable&lt;VerticalTextAlignmentValues&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **IndentationBefore**

```csharp
public Nullable<int> IndentationBefore { get; set; }
```

#### Property Value

[Nullable&lt;Int32&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **IndentationAfter**

```csharp
public Nullable<int> IndentationAfter { get; set; }
```

#### Property Value

[Nullable&lt;Int32&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **IndentationFirstLine**

```csharp
public Nullable<int> IndentationFirstLine { get; set; }
```

#### Property Value

[Nullable&lt;Int32&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **IndentationHanging**

```csharp
public Nullable<int> IndentationHanging { get; set; }
```

#### Property Value

[Nullable&lt;Int32&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **TextDirection**

```csharp
public Nullable<TextDirectionValues> TextDirection { get; set; }
```

#### Property Value

[Nullable&lt;TextDirectionValues&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **LineSpacingRule**

```csharp
public Nullable<LineSpacingRuleValues> LineSpacingRule { get; set; }
```

#### Property Value

[Nullable&lt;LineSpacingRuleValues&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **LineSpacing**

```csharp
public Nullable<int> LineSpacing { get; set; }
```

#### Property Value

[Nullable&lt;Int32&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **LineSpacingBefore**

```csharp
public Nullable<int> LineSpacingBefore { get; set; }
```

#### Property Value

[Nullable&lt;Int32&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **LineSpacingAfter**

```csharp
public Nullable<int> LineSpacingAfter { get; set; }
```

#### Property Value

[Nullable&lt;Int32&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **IsEmpty**

```csharp
public bool IsEmpty { get; }
```

#### Property Value

[Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

### **IsPageBreak**

```csharp
public bool IsPageBreak { get; }
```

#### Property Value

[Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

### **IsBreak**

```csharp
public bool IsBreak { get; }
```

#### Property Value

[Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

### **Bold**

```csharp
public bool Bold { get; set; }
```

#### Property Value

[Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

### **Italic**

```csharp
public bool Italic { get; set; }
```

#### Property Value

[Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

### **Underline**

```csharp
public Nullable<UnderlineValues> Underline { get; set; }
```

#### Property Value

[Nullable&lt;UnderlineValues&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **DoNotCheckSpellingOrGrammar**

```csharp
public bool DoNotCheckSpellingOrGrammar { get; set; }
```

#### Property Value

[Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

### **Spacing**

```csharp
public Nullable<int> Spacing { get; set; }
```

#### Property Value

[Nullable&lt;Int32&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **Strike**

```csharp
public bool Strike { get; set; }
```

#### Property Value

[Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

### **DoubleStrike**

```csharp
public bool DoubleStrike { get; set; }
```

#### Property Value

[Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

### **FontSize**

```csharp
public Nullable<int> FontSize { get; set; }
```

#### Property Value

[Nullable&lt;Int32&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **Color**

```csharp
public Color Color { get; set; }
```

#### Property Value

Color<br>

### **ColorHex**

```csharp
public string ColorHex { get; set; }
```

#### Property Value

[String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

### **ThemeColor**

```csharp
public Nullable<ThemeColorValues> ThemeColor { get; set; }
```

#### Property Value

[Nullable&lt;ThemeColorValues&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **Highlight**

```csharp
public Nullable<HighlightColorValues> Highlight { get; set; }
```

#### Property Value

[Nullable&lt;HighlightColorValues&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **CapsStyle**

```csharp
public CapsStyle CapsStyle { get; set; }
```

#### Property Value

[CapsStyle](./officeimo.word.capsstyle.md)<br>

### **FontFamily**

```csharp
public string FontFamily { get; set; }
```

#### Property Value

[String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

## Constructors

### **WordParagraph(WordDocument, Boolean, Boolean)**

```csharp
public WordParagraph(WordDocument document, bool newParagraph, bool newRun)
```

#### Parameters

`document` [WordDocument](./officeimo.word.worddocument.md)<br>

`newParagraph` [Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

`newRun` [Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

### **WordParagraph(WordDocument, Paragraph)**

```csharp
public WordParagraph(WordDocument document, Paragraph paragraph)
```

#### Parameters

`document` [WordDocument](./officeimo.word.worddocument.md)<br>

`paragraph` Paragraph<br>

### **WordParagraph(WordDocument, Paragraph, Run)**

```csharp
public WordParagraph(WordDocument document, Paragraph paragraph, Run run)
```

#### Parameters

`document` [WordDocument](./officeimo.word.worddocument.md)<br>

`paragraph` Paragraph<br>

`run` Run<br>

## Methods

### **SetBold(Boolean)**

```csharp
public WordParagraph SetBold(bool isBold)
```

#### Parameters

`isBold` [Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

#### Returns

[WordParagraph](./officeimo.word.wordparagraph.md)<br>

### **SetItalic(Boolean)**

```csharp
public WordParagraph SetItalic(bool isItalic)
```

#### Parameters

`isItalic` [Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

#### Returns

[WordParagraph](./officeimo.word.wordparagraph.md)<br>

### **SetUnderline(UnderlineValues)**

```csharp
public WordParagraph SetUnderline(UnderlineValues underline)
```

#### Parameters

`underline` UnderlineValues<br>

#### Returns

[WordParagraph](./officeimo.word.wordparagraph.md)<br>

### **SetSpacing(Int32)**

```csharp
public WordParagraph SetSpacing(int spacing)
```

#### Parameters

`spacing` [Int32](https://docs.microsoft.com/en-us/dotnet/api/system.int32)<br>

#### Returns

[WordParagraph](./officeimo.word.wordparagraph.md)<br>

### **SetStrike(Boolean)**

```csharp
public WordParagraph SetStrike(bool isStrike)
```

#### Parameters

`isStrike` [Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

#### Returns

[WordParagraph](./officeimo.word.wordparagraph.md)<br>

### **SetDoubleStrike(Boolean)**

```csharp
public WordParagraph SetDoubleStrike(bool isDoubleStrike)
```

#### Parameters

`isDoubleStrike` [Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

#### Returns

[WordParagraph](./officeimo.word.wordparagraph.md)<br>

### **SetFontSize(Int32)**

```csharp
public WordParagraph SetFontSize(int fontSize)
```

#### Parameters

`fontSize` [Int32](https://docs.microsoft.com/en-us/dotnet/api/system.int32)<br>

#### Returns

[WordParagraph](./officeimo.word.wordparagraph.md)<br>

### **SetFontFamily(String)**

```csharp
public WordParagraph SetFontFamily(string fontFamily)
```

#### Parameters

`fontFamily` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

#### Returns

[WordParagraph](./officeimo.word.wordparagraph.md)<br>

### **SetColorHex(String)**

```csharp
public WordParagraph SetColorHex(string color)
```

#### Parameters

`color` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

#### Returns

[WordParagraph](./officeimo.word.wordparagraph.md)<br>

### **SetColor(Color)**

```csharp
public WordParagraph SetColor(Color color)
```

#### Parameters

`color` Color<br>

#### Returns

[WordParagraph](./officeimo.word.wordparagraph.md)<br>

### **SetHighlight(HighlightColorValues)**

```csharp
public WordParagraph SetHighlight(HighlightColorValues highlight)
```

#### Parameters

`highlight` HighlightColorValues<br>

#### Returns

[WordParagraph](./officeimo.word.wordparagraph.md)<br>

### **SetCapsStyle(CapsStyle)**

```csharp
public WordParagraph SetCapsStyle(CapsStyle capsStyle)
```

#### Parameters

`capsStyle` [CapsStyle](./officeimo.word.capsstyle.md)<br>

#### Returns

[WordParagraph](./officeimo.word.wordparagraph.md)<br>

### **SetText(String)**

```csharp
public WordParagraph SetText(string text)
```

#### Parameters

`text` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

#### Returns

[WordParagraph](./officeimo.word.wordparagraph.md)<br>

### **SetStyle(WordParagraphStyles)**

```csharp
public WordParagraph SetStyle(WordParagraphStyles style)
```

#### Parameters

`style` [WordParagraphStyles](./officeimo.word.wordparagraphstyles.md)<br>

#### Returns

[WordParagraph](./officeimo.word.wordparagraph.md)<br>

### **CombineIdenticalRuns(Body)**

Combines the identical runs.

```csharp
public static void CombineIdenticalRuns(Body body)
```

#### Parameters

`body` Body<br>

### **AddText(String)**

Add a text to existing paragraph

```csharp
public WordParagraph AddText(string text)
```

#### Parameters

`text` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

#### Returns

[WordParagraph](./officeimo.word.wordparagraph.md)<br>

### **AddImage(String, Nullable&lt;Double&gt;, Nullable&lt;Double&gt;)**

```csharp
public WordParagraph AddImage(string filePathImage, Nullable<double> width, Nullable<double> height)
```

#### Parameters

`filePathImage` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

`width` [Nullable&lt;Double&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

`height` [Nullable&lt;Double&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

#### Returns

[WordParagraph](./officeimo.word.wordparagraph.md)<br>

### **AddImage(String)**

```csharp
public WordParagraph AddImage(string filePathImage)
```

#### Parameters

`filePathImage` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

#### Returns

[WordParagraph](./officeimo.word.wordparagraph.md)<br>

### **AddBreak(Nullable&lt;BreakValues&gt;)**

Add Break to the paragraph. By default it adds soft break (SHIFT+ENTER)

```csharp
public WordParagraph AddBreak(Nullable<BreakValues> breakType)
```

#### Parameters

`breakType` [Nullable&lt;BreakValues&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

#### Returns

[WordParagraph](./officeimo.word.wordparagraph.md)<br>

### **Remove()**

Remove the paragraph from WordDocument

```csharp
public void Remove()
```

#### Exceptions

[InvalidOperationException](https://docs.microsoft.com/en-us/dotnet/api/system.invalidoperationexception)<br>

### **AddParagraphAfterSelf()**

```csharp
public WordParagraph AddParagraphAfterSelf()
```

#### Returns

[WordParagraph](./officeimo.word.wordparagraph.md)<br>

### **AddParagraphAfterSelf(WordSection)**

```csharp
public WordParagraph AddParagraphAfterSelf(WordSection section)
```

#### Parameters

`section` [WordSection](./officeimo.word.wordsection.md)<br>

#### Returns

[WordParagraph](./officeimo.word.wordparagraph.md)<br>

### **AddParagraphBeforeSelf()**

```csharp
public WordParagraph AddParagraphBeforeSelf()
```

#### Returns

[WordParagraph](./officeimo.word.wordparagraph.md)<br>

### **AddComment(String, String, String)**

Add a comment to paragraph

```csharp
public void AddComment(string author, string initials, string comment)
```

#### Parameters

`author` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

`initials` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

`comment` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

### **AddHorizontalLine(BorderValues, Nullable&lt;Color&gt;, UInt32, UInt32)**

Add horizontal line (sometimes known as horizontal rule) to document

```csharp
public WordParagraph AddHorizontalLine(BorderValues lineType, Nullable<Color> color, uint size, uint space)
```

#### Parameters

`lineType` BorderValues<br>

`color` [Nullable&lt;Color&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

`size` [UInt32](https://docs.microsoft.com/en-us/dotnet/api/system.uint32)<br>

`space` [UInt32](https://docs.microsoft.com/en-us/dotnet/api/system.uint32)<br>

#### Returns

[WordParagraph](./officeimo.word.wordparagraph.md)<br>

### **AddBookmark(String)**

```csharp
public WordParagraph AddBookmark(string bookmarkName)
```

#### Parameters

`bookmarkName` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

#### Returns

[WordParagraph](./officeimo.word.wordparagraph.md)<br>

### **AddField(WordFieldType, Nullable&lt;WordFieldFormat&gt;, Boolean)**

```csharp
public WordParagraph AddField(WordFieldType wordFieldType, Nullable<WordFieldFormat> wordFieldFormat, bool advanced)
```

#### Parameters

`wordFieldType` [WordFieldType](./officeimo.word.wordfieldtype.md)<br>

`wordFieldFormat` [Nullable&lt;WordFieldFormat&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

`advanced` [Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

#### Returns

[WordParagraph](./officeimo.word.wordparagraph.md)<br>

### **AddHyperLink(String, Uri, Boolean, String, Boolean)**

```csharp
public WordParagraph AddHyperLink(string text, Uri uri, bool addStyle, string tooltip, bool history)
```

#### Parameters

`text` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

`uri` Uri<br>

`addStyle` [Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

`tooltip` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

`history` [Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

#### Returns

[WordParagraph](./officeimo.word.wordparagraph.md)<br>

### **AddHyperLink(String, String, Boolean, String, Boolean)**

```csharp
public WordParagraph AddHyperLink(string text, string anchor, bool addStyle, string tooltip, bool history)
```

#### Parameters

`text` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

`anchor` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

`addStyle` [Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

`tooltip` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

`history` [Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

#### Returns

[WordParagraph](./officeimo.word.wordparagraph.md)<br>
