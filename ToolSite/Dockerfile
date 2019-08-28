FROM microsoft/dotnet:2.1-aspnetcore-runtime AS base
WORKDIR /app
EXPOSE 80

FROM microsoft/dotnet:2.1-sdk AS build
WORKDIR /src
COPY ToolSite/ToolSite.csproj ToolSite/
RUN dotnet restore ToolSite/ToolSite.csproj
COPY . .
# WORKDIR /src/ToolSite
# RUN dotnet build ToolSite.csproj -c Release -o /app

FROM build AS publish
WORKDIR /src/ToolSite
RUN dotnet publish ToolSite.csproj -c Release -o /app

FROM base AS final
WORKDIR /app
COPY --from=publish /app .
ENTRYPOINT ["dotnet", "ToolSite.dll"]
