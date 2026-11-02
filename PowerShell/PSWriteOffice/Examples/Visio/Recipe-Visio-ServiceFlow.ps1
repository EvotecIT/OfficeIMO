$path = '.\Service-Flow.vsdx'

New-OfficeVisio -Path $path -Width 10in -Height 6in -Title 'Service request flow' {
    VisioTextBox 'Service request flow' -X 5in -Y 5.4in -Width 4in -Height 0.4in -FillColor '#FFFFFF' -LineColor '#FFFFFF'
    VisioRectangle -Key request -Text 'Request' -X 2in -Y 3in -Width 1.8in -Height 0.9in -FillColor '#DBEAFE'
    VisioRectangle -Key approve -Text 'Approval' -X 5in -Y 3in -Width 1.8in -Height 0.9in -FillColor '#FEF3C7'
    VisioRectangle -Key deliver -Text 'Delivery' -X 8in -Y 3in -Width 1.8in -Height 0.9in -FillColor '#DCFCE7'
    VisioConnector -From request -To approve -FromSide Right -ToSide Left -EndArrow Triangle
    VisioConnector -From approve -To deliver -FromSide Right -ToSide Left -EndArrow Triangle
}
