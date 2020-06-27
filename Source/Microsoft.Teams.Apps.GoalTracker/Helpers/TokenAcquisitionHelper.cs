// <copyright file="TokenAcquisitionHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Identity.Client;

    /// <summary>
    /// Gets and sets access token in cache.
    /// </summary>
    public class TokenAcquisitionHelper
    {
        /// <summary>
        /// Instance of confidential client app to access web API.
        /// </summary>
        private readonly IConfidentialClientApplication confidentialClientApp;

        /// <summary>
        /// Represents scopes required by MsalNet for accessing token.
        /// </summary>
        private readonly string[] scopesRequestedByMsalNet = new string[]
        {
            "openid",
            "profile",
            "offline_access",
        };

        /// <summary>
        /// Initializes a new instance of the <see cref="TokenAcquisitionHelper"/> class.
        /// </summary>
        /// <param name="confidentialClientApp">Instance of ConfidentialClientApplication class.</param>
        public TokenAcquisitionHelper(
            IConfidentialClientApplication confidentialClientApp)
        {
            this.confidentialClientApp = confidentialClientApp;
        }

        /// <summary>
        /// Adds token to cache.
        /// </summary>
        /// <param name="graphScopes">Graph scopes to be added to token.</param>
        /// <param name="jwtToken">JWT bearer token.</param>
        /// <returns>Token with graph scopes.</returns>
        public async Task<string> AddTokenToCacheFromJwtAsync(string graphScopes, string jwtToken)
        {
            graphScopes = graphScopes ?? throw new ArgumentNullException(nameof(graphScopes));
            jwtToken = jwtToken ?? throw new ArgumentNullException(nameof(jwtToken));
            UserAssertion userAssertion = new UserAssertion(jwtToken, "urn:ietf:params:oauth:grant-type:jwt-bearer");
            IEnumerable<string> requestedScopes = graphScopes.Split(new char[] { ' ' }, System.StringSplitOptions.RemoveEmptyEntries).ToList();

            // Result to make sure that the cache is filled-in before the controller tries to get access tokens
            var result = await this.confidentialClientApp.AcquireTokenOnBehalfOf(
                requestedScopes.Except(this.scopesRequestedByMsalNet),
                userAssertion)
                .ExecuteAsync();
            return result.AccessToken;
        }
    }
}
