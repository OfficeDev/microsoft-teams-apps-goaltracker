// <copyright file="ServicesExtension.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using Microsoft.AspNetCore.Builder;
    using Microsoft.AspNetCore.Localization;
    using Microsoft.Azure.Search;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Azure;
    using Microsoft.Bot.Builder.BotFramework;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Extensions.DependencyInjection;
    using Microsoft.Identity.Client;
    using Microsoft.Teams.Apps.GoalTracker.Common;
    using Microsoft.Teams.Apps.GoalTracker.Helpers;
    using Microsoft.Teams.Apps.GoalTracker.Models;

    /// <summary>
    /// Class to extend ServiceCollection.
    /// </summary>
    public static class ServicesExtension
    {
        /// <summary>
        /// Azure Search service index name for personal goal details.
        /// </summary>
        private const string PersonalGoalIndexName = "personal-goal-index";

        /// <summary>
        /// Adds application configuration settings to specified IServiceCollection.
        /// </summary>
        /// <param name="services">Collection of services.</param>
        /// <param name="configuration">Application configuration properties.</param>
        public static void AddConfigurationSettings(this IServiceCollection services, IConfiguration configuration)
        {
            services.Configure<GoalTrackerActivityHandlerOptions>(options =>
            {
                options.TenantId = configuration.GetValue<string>("Bot:TenantId");
                options.AppBaseUri = configuration.GetValue<string>("Bot:AppBaseUri");
                options.GoalsTabEntityId = configuration.GetValue<string>("Bot:GoalsTabEntityId");
                options.ManifestId = configuration.GetValue<string>("Bot:ManifestId");
            });

            services.Configure<TokenOptions>(options =>
            {
                options.SecurityKey = configuration.GetValue<string>("Token:SecurityKey");
            });

            services.Configure<StorageOptions>(options =>
            {
                options.ConnectionString = configuration.GetValue<string>("Storage:ConnectionString");
            });

            services.Configure<TelemetryOptions>(options =>
            {
                options.InstrumentationKey = configuration.GetValue<string>("ApplicationInsights:InstrumentationKey");
            });

            services.Configure<SearchServiceOptions>(options =>
            {
                options.SearchServiceName = configuration.GetValue<string>("Search:SearchServiceName");
                options.SearchServiceQueryApiKey = configuration.GetValue<string>("Search:SearchServiceQueryApiKey");
                options.SearchServiceAdminApiKey = configuration.GetValue<string>("Search:SearchServiceAdminApiKey");
            });

            services.Configure<MicrosoftAppOptions>(options =>
            {
                options.ClientId = configuration.GetValue<string>("MicrosoftAppId");
                options.ClientSecret = configuration.GetValue<string>("MicrosoftAppPassword");
                options.TenantId = configuration.GetValue<string>("TenantId");
            });

            services.Configure<AzureAdOptions>(options =>
            {
                options.ClientId = configuration.GetValue<string>("AzureAd:ClientId");
                options.ClientSecret = configuration.GetValue<string>("AzureAd:ClientSecret");
                options.GraphScope = configuration.GetValue<string>("AzureAd:GraphScope");
                options.TenantId = configuration.GetValue<string>("AzureAd:TenantId");
            });
        }

        /// <summary>
        /// Adds providers to specified IServiceCollection.
        /// </summary>
        /// <param name="services">Collection of services.</param>
        public static void AddProviders(this IServiceCollection services)
        {
            services
                .AddTransient<IPersonalGoalStorageProvider, PersonalGoalStorageProvider>();
            services
                .AddTransient<IPersonalGoalNoteStorageProvider, PersonalGoalNoteStorageProvider>();
            services
                .AddTransient<ITeamGoalStorageProvider, TeamGoalStorageProvider>();
            services
                .AddSingleton<IPersonalGoalSearchService, PersonalGoalSearchService>();
            services
                .AddSingleton<ITeamStorageProvider, TeamStorageProvider>();
            services
                .AddHostedService<GoalReminderNotificationService>();
            services
                .AddHostedService<GoalDeletionBackgroundService>();
            services
                .AddHostedService<GoalBackgroundService>();
            services
                .AddSingleton<BackgroundTaskWrapper>();
        }

        /// <summary>
        /// Adds helpers to specified IServiceCollection.
        /// </summary>
        /// <param name="services">Collection of services.</param>
        /// <param name="configuration">Application configuration properties.</param>
        public static void AddHelpers(this IServiceCollection services, IConfiguration configuration)
        {
            services
                .AddApplicationInsightsTelemetry(configuration.GetValue<string>("ApplicationInsights:InstrumentationKey"));
            services
                .AddSingleton<IGoalReminderActivityHelper, GoalReminderActivityHelper>();
            services
                .AddSingleton<CardHelper>();
            services
                .AddSingleton<ActivityHelper>();
            services
                .AddSingleton<GoalHelper>();
            services
                .AddSingleton<TokenAcquisitionHelper>();
            services
                .AddSingleton<ITeamsInfoHelper, TeamsInfoHelper>();
#pragma warning disable CA2000 // Dispose objects before losing scope - PersonalGoalSearchService uses IDisposable where search service client is injected.
            services
                .AddSingleton(new SearchServiceClient(configuration.GetValue<string>("Search:SearchServiceName"), new SearchCredentials(configuration.GetValue<string>("Search:SearchServiceAdminApiKey"))));
            services
                .AddSingleton(new SearchIndexClient(configuration.GetValue<string>("Search:SearchServiceName"), PersonalGoalIndexName, new SearchCredentials(configuration.GetValue<string>("Search:SearchServiceQueryApiKey"))));
#pragma warning restore CA2000 // Dispose objects before losing scope
        }

        /// <summary>
        /// Adds credential providers for authentication.
        /// </summary>
        /// <param name="services">Collection of services.</param>
        /// <param name="configuration">Application configuration properties.</param>
        public static void AddCredentialProviders(this IServiceCollection services, IConfiguration configuration)
        {
            services
                .AddSingleton<ICredentialProvider, ConfigurationCredentialProvider>();
            services
                .AddSingleton(new MicrosoftAppCredentials(configuration.GetValue<string>("MicrosoftAppId"), configuration.GetValue<string>("MicrosoftAppPassword")));
        }

        /// <summary>
        /// Add confidential credential provider to access api.
        /// </summary>
        /// <param name="services">Collection of services.</param>
        /// <param name="configuration">Application configuration properties.</param>
        public static void AddConfidentialCredentialProvider(this IServiceCollection services, IConfiguration configuration)
        {
            configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));

            IConfidentialClientApplication confidentialClientApp = ConfidentialClientApplicationBuilder.Create(configuration["AzureAd:ClientId"])
                .WithClientSecret(configuration["AzureAd:ClientSecret"])
                .Build();
            services.AddSingleton<IConfidentialClientApplication>(confidentialClientApp);
        }

        /// <summary>
        /// Adds localization settings to specified IServiceCollection.
        /// </summary>
        /// <param name="services">Collection of services.</param>
        /// <param name="configuration">Application configuration properties.</param>
        public static void AddLocalizationSettings(this IServiceCollection services, IConfiguration configuration)
        {
            services.AddLocalization(options => options.ResourcesPath = "Resources");
            services.Configure<RequestLocalizationOptions>(options =>
            {
                var defaultCulture = CultureInfo.GetCultureInfo(configuration.GetValue<string>("i18n:DefaultCulture"));
                var supportedCultures = configuration.GetValue<string>("i18n:SupportedCultures").Split(',')
                    .Select(culture => CultureInfo.GetCultureInfo(culture))
                    .ToList();

                options.DefaultRequestCulture = new RequestCulture(defaultCulture);
                options.SupportedCultures = supportedCultures;
                options.SupportedUICultures = supportedCultures;

                options.RequestCultureProviders = new List<IRequestCultureProvider>
                {
                    new BotLocalizationCultureProvider(),
                };
            });
        }

        /// <summary>
        /// Adds user state and conversation state to specified IServiceCollection.
        /// </summary>
        /// <param name="services">Collection of services.</param>
        /// <param name="configuration">Application configuration properties.</param>
        public static void AddBotStates(this IServiceCollection services, IConfiguration configuration)
        {
            // Create the User state. (Used in this bot's Dialog implementation.)
            services.AddSingleton<UserState>();

            // Create the Conversation state. (Used by the Dialog system itself.)
            services.AddSingleton<ConversationState>();

            // For conversation state.
            services.AddSingleton<IStorage>(new AzureBlobStorage(configuration.GetValue<string>("Storage:ConnectionString"), "bot-state"));
        }
    }
}