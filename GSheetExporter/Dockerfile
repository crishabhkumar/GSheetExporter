# Stage 1: Build the app with .NET 9 SDK (Preview)
FROM mcr.microsoft.com/dotnet/sdk:9.0-preview AS build
WORKDIR /app

# Copy csproj and restore dependencies
COPY *.csproj ./
RUN dotnet restore

# Copy the rest of the project files
COPY . ./

# Publish the app to /out folder
RUN dotnet publish -c Release -o /out

# Stage 2: Build a runtime-only image using .NET 9 ASP.NET runtime
FROM mcr.microsoft.com/dotnet/aspnet:9.0-preview AS runtime
WORKDIR /app

COPY --from=build /out ./

EXPOSE 80
ENTRYPOINT ["dotnet", "GSheetExporter.dll"]
