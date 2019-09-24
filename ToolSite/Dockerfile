FROM microsoft/dotnet:2.2-aspnetcore-runtime AS base
WORKDIR /app
# RUN apt-get update \
#     && apt-get install -y --allow-unauthenticated \
#    		libgdiplus \
#      && rm -rf /var/lib/apt/lists/*
EXPOSE 80

FROM microsoft/dotnet:2.2-sdk AS build
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