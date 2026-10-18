$path = '.\Service-Flow.vsdx'

OfficeVisio -Path $path -Width 10 -Height 6 -Title 'Service request flow' {
    VisioTextBox 'Service request flow' -X 5 -Y 5.4 -Width 4 -Height 0.4 -FillColor '#FFFFFF' -LineColor '#FFFFFF'
    VisioRectangle -Key request -Text 'Request' -X 2 -Y 3 -Width 1.8 -Height 0.9 -FillColor '#DBEAFE'
    VisioRectangle -Key approve -Text 'Approval' -X 5 -Y 3 -Width 1.8 -Height 0.9 -FillColor '#FEF3C7'
    VisioRectangle -Key deliver -Text 'Delivery' -X 8 -Y 3 -Width 1.8 -Height 0.9 -FillColor '#DCFCE7'
    VisioConnector -From request -To approve -FromSide Right -ToSide Left -EndArrow Triangle
    VisioConnector -From approve -To deliver -FromSide Right -ToSide Left -EndArrow Triangle
}
