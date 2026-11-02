# Each top-level heading or --- starts a slide: headings become titles, lists become bullets,
# tables flow into place and blockquotes become speaker notes.
@'
# Quarterly business review
Performance, decisions and next-quarter priorities

## Performance trend
- Revenue grew 36% from Q1 to Q4
- Operating margin improved five percentage points
- Renewals held above target in every region

> Explain the margin improvement before discussing revenue.

## Next-quarter priorities
| Priority | Owner | Target |
| --- | --- | --- |
| Customer onboarding | Product | Reduce setup time by 25% |
| Renewal risk | Sales | Review top 20 accounts |
| Delivery capacity | Operations | Add two automation lanes |

> Close by assigning each priority to a named owner.
'@ | Set-Content -Path .\Review.md

Convert-OfficeDocument -Path .\Review.md -OutputPath .\Review.pptx -Show
