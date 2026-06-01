FROM mcr.microsoft.com/dotnet/framework/sdk:4.8-windowsservercore-ltsc2022 AS build

WORKDIR /app

COPY . .

RUN msbuild ComputerClubInventory/ComputerClubInventory.csproj /p:Configuration=Release

FROM mcr.microsoft.com/dotnet/framework/runtime:4.8-windowsservercore-ltsc2022

WORKDIR /app

COPY --from=build /app/ComputerClubInventory/bin/Release/ .

CMD ["ComputerClubInventory.exe"]