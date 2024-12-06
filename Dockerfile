# Etapa de construcción
FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /source

# Copiar archivos del proyecto y restaurar dependencias
COPY *.sln .
COPY api-excel-new.csproj .
RUN dotnet restore

# Copiar el resto del código fuente y compilar
COPY . .
RUN dotnet publish -c Release -o /app

# Etapa final
FROM mcr.microsoft.com/dotnet/aspnet:8.0
WORKDIR /app
COPY --from=build /app ./

# Exponer el puerto correcto
EXPOSE 8001
ENTRYPOINT ["dotnet", "api-excel-new.dll"]
