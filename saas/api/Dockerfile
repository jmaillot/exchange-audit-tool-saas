FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /src
# Build context is saas/api (see build: ./api in docker-compose files).
COPY . .
RUN dotnet publish -c Release -o /app

FROM mcr.microsoft.com/dotnet/aspnet:8.0
WORKDIR /app
COPY --from=build /app .
ENV ASPNETCORE_URLS=http://+:8080 EAT_DATA_DIR=/data EAT_WORKER_URL=http://worker:8081
VOLUME /data
EXPOSE 8080
ENTRYPOINT ["dotnet", "ExchangeAuditSaaS.dll"]
