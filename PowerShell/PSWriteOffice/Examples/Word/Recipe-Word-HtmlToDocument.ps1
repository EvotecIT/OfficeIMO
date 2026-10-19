$html = @'
<h1>Service review</h1>
<p>The weekly review is <strong>ready</strong>.</p>
<ul><li>Identity is healthy.</li><li>Messaging needs follow-up.</li></ul>
<table><tr><th>Area</th><th>Owner</th></tr><tr><td>Identity</td><td>Platform</td></tr></table>
'@

ConvertFrom-OfficeWordHtml -Html $html -OutputPath '.\Service-Review.docx'
ConvertTo-OfficeWordMarkdown -Path '.\Service-Review.docx' -OutputPath '.\Service-Review.md'
