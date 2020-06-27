// <copyright file="AuthenticationServiceCollectionExtensions.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Authentication
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Authentication.AzureAD.UI;
    using Microsoft.AspNetCore.Authentication.JwtBearer;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Extensions.DependencyInjection;
    using Microsoft.IdentityModel.Tokens;
    using Microsoft.Teams.Apps.GoalTracker.Authentication.AuthenticationPolicy;
    using Microsoft.Teams.Apps.GoalTracker.Helpers;

    /// <summary>
    /// Extension class for registering authentication services in DI container.
    /// </summary>
    public static class AuthenticationServiceCollectionExtensions
    {
        private const string ClientIdConfigurationSettingsKey = "AzureAd:ClientId";
        private const string TenantIdConfigurationSettingsKey = "AzureAd:TenantId";
        private const string ApplicationIdURIConfigurationSettingsKey = "AzureAd:ApplicationIdURI";
        private const string ValidIssuersConfigurationSettingsKey = "AzureAd:ValidIssuers";
        private const string GraphScopeConfigurationSettingsKey = "AzureAd:GraphScope";

        /// <summary>
        /// Extension method to register the authentication services.
        /// </summary>
        /// <param name="services">IServiceCollection instance.</param>
        /// <param name="configuration">IConfiguration instance.</param>
        public static void AddGoalTrackerAuthentication(this IServiceCollection services, IConfiguration configuration)
        {
            RegisterAuthenticationServices(services, configuration);
            RegisterAuthorizationPolicy(services);
        }

        // This method works specifically for single tenant application.
        private static void RegisterAuthenticationServices(
            IServiceCollection services,
            IConfiguration configuration)
        {
            ValidateAuthenticationConfigurationSettings(configuration);

            services.AddAuthentication(options => { options.DefaultScheme = JwtBearerDefaults.AuthenticationScheme; })
                .AddJwtBearer(options =>
                {
                    var azureADOptions = new AzureADOptions();
                    configuration.Bind("AzureAd", azureADOptions);
                    options.Authority = $"{azureADOptions.Instance}{azureADOptions.TenantId}/v2.0";
                    options.TokenValidationParameters = new TokenValidationParameters
                    {
                        ValidAudiences = GetValidAudiences(configuration),
                        ValidIssuers = GetValidIssuers(configuration),
                        AudienceValidator = AudienceValidator,
                    };
                    options.Events = new JwtBearerEvents
                    {
                        OnTokenValidated = async context =>
                        {
                            var tokenAcquisition = context.HttpContext.RequestServices.GetRequiredService<TokenAcquisitionHelper>();
                            context.Success();

                            // Adds the token to the cache, and also handles the incremental consent and claim challenges
                            await tokenAcquisition.AddTokenToCacheFromJwtAsync(configuration[AuthenticationServiceCollectionExtensions.GraphScopeConfigurationSettingsKey], context.Request.Headers["Authorization"].ToString().Replace("Bearer", string.Empty, StringComparison.InvariantCulture));
                            await Task.FromResult(0);
                        },
                    };
                });
        }

        /// <summary>
        /// Validates authentication configuration settings provided in app settings.
        /// </summary>
        /// <param name="configuration">Application settings.</param>
        private static void ValidateAuthenticationConfigurationSettings(IConfiguration configuration)
        {
            var clientId = configuration[ClientIdConfigurationSettingsKey];
            if (string.IsNullOrWhiteSpace(clientId))
            {
                throw new ApplicationException("AzureAD ClientId is missing in the configuration file.");
            }

            var tenantId = configuration[TenantIdConfigurationSettingsKey];
            if (string.IsNullOrWhiteSpace(tenantId))
            {
                throw new ApplicationException("AzureAD TenantId is missing in the configuration file.");
            }

            var applicationIdURI = configuration[ApplicationIdURIConfigurationSettingsKey];
            if (string.IsNullOrWhiteSpace(applicationIdURI))
            {
                throw new ApplicationException("AzureAD ApplicationIdURI is missing in the configuration file.");
            }

            var validIssuers = configuration[ValidIssuersConfigurationSettingsKey];
            if (string.IsNullOrWhiteSpace(validIssuers))
            {
                throw new ApplicationException("AzureAD ValidIssuers is missing in the configuration file.");
            }
        }

        /// <summary>
        /// Get application settings for given key.
        /// </summary>
        /// <param name="configuration">Application settings.</param>
        /// <param name="configurationSettingsKey">Settings key.</param>
        /// <returns>Returns value associated with the key provided.</returns>
        private static IEnumerable<string> GetSettings(IConfiguration configuration, string configurationSettingsKey)
        {
            var configurationSettingsValue = configuration[configurationSettingsKey];
            var settings = configurationSettingsValue
                ?.Split(new char[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries)
                ?.Select(settingValue => settingValue.Trim());
            if (settings == null)
            {
                throw new ApplicationException($"{configurationSettingsKey} does not contain a valid value in the configuration file.");
            }

            return settings;
        }

        /// <summary>
        /// Get valid audiences from app settings.
        /// </summary>
        /// <param name="configuration">Application settings.</param>
        /// <returns>Returns  valid audiences from app settings</returns>
        private static IEnumerable<string> GetValidAudiences(IConfiguration configuration)
        {
            var clientId = configuration[ClientIdConfigurationSettingsKey];

            var applicationIdURI = configuration[ApplicationIdURIConfigurationSettingsKey];

            var validAudiences = new List<string> { clientId, applicationIdURI.ToUpperInvariant() };

            return validAudiences;
        }

        /// <summary>
        /// Get valid issuers from app settings.
        /// </summary>
        /// <param name="configuration">Application settings.</param>
        /// <returns>Returns  valid issuers from app settings</returns>
        private static IEnumerable<string> GetValidIssuers(IConfiguration configuration)
        {
            var tenantId = configuration[TenantIdConfigurationSettingsKey];

            var validIssuers =
                GetSettings(
                    configuration,
                    ValidIssuersConfigurationSettingsKey);

            validIssuers = validIssuers.Select(validIssuer => validIssuer.Replace("TENANT_ID", tenantId, StringComparison.OrdinalIgnoreCase));

            return validIssuers;
        }

        /// <summary>
        /// Validates audience.
        /// </summary>
        /// <param name="tokenAudiences">Valid audience token.</param>
        /// <param name="securityToken">Valid security token.</param>
        /// <param name="validationParameters">Valid audiences.</param>
        /// <returns>Returns true for valid audience, else false.</returns>
        private static bool AudienceValidator(
            IEnumerable<string> tokenAudiences,
            SecurityToken securityToken,
            TokenValidationParameters validationParameters)
        {
            if (tokenAudiences == null || !tokenAudiences.Any())
            {
                throw new ApplicationException("No audience defined in token!");
            }

            var validAudiences = validationParameters.ValidAudiences;
            if (validAudiences == null || !validAudiences.Any())
            {
                throw new ApplicationException("No valid audiences defined in validationParameters!");
            }

            foreach (var tokenAudience in tokenAudiences)
            {
                if (validAudiences.Any(validAudience => validAudience.Equals(tokenAudience, StringComparison.OrdinalIgnoreCase)))
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Extension method to register the authorization policies.
        /// </summary>
        /// <param name="services">IServiceCollection instance.</param>
        private static void RegisterAuthorizationPolicy(IServiceCollection services)
        {
            services.AddAuthorization(options =>
            {
                var mustContainValidUserRequirement = new MustBeTeamMemberRequirement();
                options.AddPolicy(
                    PolicyNames.MustBePartOfTeamPolicy,
                    policyBuilder => policyBuilder.AddRequirements(mustContainValidUserRequirement));
            });

            services.AddSingleton<IAuthorizationHandler, MustBePartOfTeamHandler>();
        }
    }
}