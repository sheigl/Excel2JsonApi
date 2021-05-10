############################################################
## DOTNET                                                 ##
############################################################

FROM mcr.microsoft.com/dotnet/sdk:5.0 AS build
WORKDIR /app

COPY ./Excel2JsonApi.csproj ./Excel2JsonApi.csproj
RUN dotnet restore ./Excel2JsonApi.csproj

# copy everything else and build app
COPY . .
RUN dotnet publish ./Excel2JsonApi.csproj -c Release -o /app/out

FROM mcr.microsoft.com/dotnet/aspnet:5.0 AS runtime
WORKDIR /app
COPY --from=build /app/out ./

ARG ENVIRONMENT=Production
ARG RESUB_PATH=/data/uploads/resub
ARG REPROCESS_PATH=/data/uploads/reprocess
ARG PROCESS_PATH=/data/uploads/process

ENV ASPNETCORE_ENVIRONMENT=${ENVIRONMENT}
ENV ASPNETCORE_ResubPath=${RESUB_PATH}
ENV ASPNETCORE_ReprocessPath=${REPROCESS_PATH}
ENV ASPNETCORE_ProcessPath=${PROCESS_PATH}

ENTRYPOINT ["dotnet", "Excel2JsonApi.dll"]