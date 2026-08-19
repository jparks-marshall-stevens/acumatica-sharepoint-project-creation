using Azure.Storage.Blobs;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using ProjectSync;
using ProjectSync.Acumatica;
using ProjectSync.Notifications;
using ProjectSync.HubSpot;
using ProjectSync.Options;
using ProjectSync.SharePoint;
using ProjectSync.State;

var host = new HostBuilder()
    .ConfigureFunctionsWorkerDefaults()
    .ConfigureAppConfiguration(config =>
    {
        config.AddEnvironmentVariables();
    })
    .ConfigureServices((context, services) =>
    {
        var configuration = context.Configuration;

        services.AddApplicationInsightsTelemetryWorkerService();
        services.ConfigureFunctionsApplicationInsights();

        // Bind strongly-typed options from configuration sections.
        services.Configure<AcumaticaOptions>(configuration.GetSection(AcumaticaOptions.SectionName));
        services.Configure<SharePointOptions>(configuration.GetSection(SharePointOptions.SectionName));
        services.Configure<StateOptions>(configuration.GetSection(StateOptions.SectionName));
        services.Configure<HubSpotOptions>(configuration.GetSection(HubSpotOptions.SectionName));
        services.Configure<NotificationOptions>(configuration.GetSection(NotificationOptions.SectionName));

        // Acumatica: token provider + GI client (typed HttpClients).
        services.AddHttpClient<AcumaticaTokenProvider>();
        services.AddHttpClient<IAcumaticaClient, AcumaticaClient>();

        // HubSpot: OAuth token provider + deals client (scoping-phase source).
        services.AddHttpClient<HubSpotTokenProvider>();
        services.AddHttpClient<IHubSpotClient, HubSpotClient>();

        // State store (Blob). Falls back to the Functions storage account when no explicit
        // State:ConnectionString is provided.
        services.AddSingleton(sp =>
        {
            var stateConn = configuration.GetSection(StateOptions.SectionName)["ConnectionString"];
            var connection = string.IsNullOrWhiteSpace(stateConn)
                ? configuration["AzureWebJobsStorage"]
                : stateConn;

            if (string.IsNullOrWhiteSpace(connection))
            {
                throw new InvalidOperationException(
                    "No storage connection string. Set State:ConnectionString or AzureWebJobsStorage.");
            }

            return new BlobServiceClient(connection);
        });
        services.AddSingleton<ILastRunStore, BlobLastRunStore>();

        // Notifications (stubbed sender for now — swap the registration for a Graph sender to go live).
        services.AddSingleton<IEmailSender, LoggingEmailSender>();
        services.AddSingleton<WorkspaceNotifier>();

        // SharePoint.
        services.AddSingleton<SharePointContextFactory>();
        services.AddSingleton<ProjectSync.SharePoint.GraphUploadLinkService>();
        services.AddSingleton<ISharePointDocumentSetService, SharePointDocumentSetService>();

        // Orchestration.
        services.AddSingleton(TimeProvider.System);
        services.AddScoped<ProjectSyncProcessor>();
        services.AddScoped<HubSpotScopingProcessor>();
    })
    .Build();

host.Run();
