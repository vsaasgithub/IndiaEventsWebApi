FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build-env
WORKDIR /app

# Build the application
COPY . ./
RUN dotnet restore
RUN dotnet publish -c Release -o out

# Final runtime image (smaller in size)
FROM mcr.microsoft.com/dotnet/aspnet:8.0
WORKDIR /app
RUN mkdir -p ./Attachments
COPY --from=build-env /app/out .
ENTRYPOINT ["dotnet", "IndiaEventsWebApi.dll"]
