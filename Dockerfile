# Alpine
FROM mcr.microsoft.com/dotnet/sdk:6.0-alpine AS build

WORKDIR /app

COPY doctrack .

# RUN dotnet publish "doctrack/doctrack.csproj" -r linux-x64 -c Release -o /app/publish \
RUN dotnet publish "doctrack/doctrack.csproj" -c Release -o /app/publish \
  --runtime alpine-x64 \
  --self-contained true \
  /p:PublishTrimmed=true \
  /p:TrimMode=Link \
  /p:PublishSingleFile=true
# RUN dotnet publish -r linux-x64 -c Release /p:PublishSingleFile=true

FROM mcr.microsoft.com/dotnet/aspnet:6.0-alpine

WORKDIR /app
COPY --from=harbor.rvision.pro/sec/mc:RELEASE.2023-05-04T18-10-16Z-scratch --chmod=755 /bin/mc /bin/mc
COPY --from=build /app/publish .
# COPY doctrack_template.docx .

# CMD ["node", "index.js"]
# CMD ["tail", "-f", "/dev/null"]