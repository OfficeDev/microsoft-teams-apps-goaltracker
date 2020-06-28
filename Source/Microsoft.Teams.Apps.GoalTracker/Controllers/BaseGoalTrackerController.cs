// <copyright file="BaseGoalTrackerController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Controllers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Net.Http.Headers;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Identity.Client;
    using Microsoft.Teams.Apps.GoalTracker.Helpers;
    using Microsoft.Teams.Apps.GoalTracker.Models;

    /// <summary>
    /// Base controller to handle API operations.
    /// </summary>
    [Route("api/[controller]")]
    [ApiController]
    public class BaseGoalTrackerController : ControllerBase
    {
        /// <summary>
        /// Instance of IOptions to read data from azure application configuration.
        /// </summary>
        private readonly IOptions<AzureAdOptions> azureAdOptions;

        /// <summary>
        /// Sends logs to the Application Insights service.
        /// </summary>
        private readonly ILogger logger;

        /// <summary>
        /// Instance of confidential client applications to access Web API.
        /// </summary>
        private readonly IConfidentialClientApplication confidentialClientApp;

        /// <summary>
        /// Instance of token acquisition helper to access token.
        /// </summary>
        private readonly TokenAcquisitionHelper tokenAcquisitionHelper;

        /// <summary>
        /// Initializes a new instance of the <see cref="BaseGoalTrackerController"/> class.
        /// </summary>
        /// <param name="confidentialClientApp">Instance of ConfidentialClientApplication class.</param>
        /// <param name="azureAdOptions">Instance of IOptions to read data from application configuration.</param>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="tokenAcquisitionHelper">Instance of token acquisition helper to access token.</param>
        public BaseGoalTrackerController(
            IConfidentialClientApplication confidentialClientApp,
            IOptions<AzureAdOptions> azureAdOptions,
            ILogger logger,
            TokenAcquisitionHelper tokenAcquisitionHelper)
        {
            this.confidentialClientApp = confidentialClientApp;
            this.azureAdOptions = azureAdOptions;
            this.logger = logger;
            this.tokenAcquisitionHelper = tokenAcquisitionHelper;
        }

        /// <summary>
        /// Gets user's Azure AD object id.
        /// </summary>
        public string UserObjectId
        {
            get
            {
                var oidClaimType = "http://schemas.microsoft.com/identity/claims/objectidentifier";
                var claim = this.User.Claims.FirstOrDefault(p => oidClaimType.Equals(p.Type, StringComparison.CurrentCulture));
                return claim.Value;
            }
        }

        /// <summary>
        /// Get user Azure AD access token.
        /// </summary>
        /// <returns>Access token with Graph scopes.</returns>
        public async Task<string> GetAccessTokenAsync()
        {
            List<string> scopeList = this.azureAdOptions.Value.GraphScope.Split(new char[] { ' ' }, System.StringSplitOptions.RemoveEmptyEntries).ToList();

            try
            {
                // Gets user account from the accounts available in token cache.
                // https://docs.microsoft.com/en-us/dotnet/api/microsoft.identity.client.clientapplicationbase.getaccountasync?view=azure-dotnet
                // Concatenation of UserObjectId and TenantId separated by a dot is used as unique identifier for getting user account.
                // https://docs.microsoft.com/en-us/dotnet/api/microsoft.identity.client.accountid.identifier?view=azure-dotnet#Microsoft_Identity_Client_AccountId_Identifier
                var account = await this.confidentialClientApp.GetAccountAsync($"{this.UserObjectId}.{this.azureAdOptions.Value.TenantId}");

                // Attempts to acquire an access token for the account from the user token cache.
                // https://docs.microsoft.com/en-us/dotnet/api/microsoft.identity.client.clientapplicationbase.acquiretokensilent?view=azure-dotnet
                AuthenticationResult result = await this.confidentialClientApp
                    .AcquireTokenSilent(scopeList, account)
                    .ExecuteAsync();
                return result.AccessToken;
            }
            catch (MsalUiRequiredException ex)
            {
                try
                {
                    // Getting new token using AddTokenToCacheFromJwtAsync as AcquireTokenSilent failed to load token from cache.
                    this.logger.LogInformation($"MSAL exception occurred and trying to acquire new token. MSAL exception details are found {ex}.");
                    var jwtToken = AuthenticationHeaderValue.Parse(this.Request.Headers["Authorization"].ToString()).Parameter;
                    return await this.tokenAcquisitionHelper.AddTokenToCacheFromJwtAsync(this.azureAdOptions.Value.GraphScope, jwtToken);
                }
                catch (MsalException msalex)
                {
                    this.logger.LogError(msalex, $"An error occurred in GetAccessTokenAsync: {msalex.Message}.");
                }

                throw;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"An error occurred in fetching token : {ex.Message}.");
                throw;
            }
        }
    }
}