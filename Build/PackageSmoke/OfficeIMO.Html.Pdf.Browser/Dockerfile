FROM mcr.microsoft.com/dotnet/runtime:10.0-noble

ENV PLAYWRIGHT_DRIVER_SEARCH_PATH=/opt/htmltinkerx/.playwright \
    PLAYWRIGHT_BROWSERS_PATH=/opt/htmltinkerx/browsers
WORKDIR /app
COPY . .
RUN HTMLTINKERX_PLAYWRIGHT_WITH_DEPS=1 dotnet OfficeIMO.Html.Pdf.Browser.PackageSmoke.dll --install-browser \
    && chown -R app:app /opt/htmltinkerx
USER app
ENTRYPOINT ["dotnet", "OfficeIMO.Html.Pdf.Browser.PackageSmoke.dll", "--capture", "/tmp/browser.pdf"]
