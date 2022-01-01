FROM mcr.microsoft.com/dotnet/aspnet:2.1 AS base
WORKDIR /app
EXPOSE 80

FROM mcr.microsoft.com/dotnet/sdk:2.1 AS build
WORKDIR /src
COPY ["ToolSite/ToolSite.csproj", "ToolSite/"]
COPY ["EpplusHelper/EpplusHelper.csproj", "EpplusHelper/"]
RUN dotnet restore "ToolSite/ToolSite.csproj"
COPY . .
#WORKDIR "/src/ToolSite"
#RUN dotnet build "ToolSite.csproj" -c Release -o /app

FROM build AS publish
WORKDIR /src/ToolSite
RUN dotnet publish "ToolSite.csproj" -c Release -o /app

FROM base AS final
WORKDIR /app
COPY --from=publish /app .
ENTRYPOINT ["dotnet", "ToolSite.dll"]