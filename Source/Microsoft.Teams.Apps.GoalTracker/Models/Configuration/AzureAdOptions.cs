// <copyright file="AzureAdOptions.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Models
{
    /// <summary>
    /// AzureAdOptions class contain value application configuration properties for Azure Active Directory.
    /// </summary>
    public class AzureAdOptions
    {
        /// <summary>
        /// Gets or sets the Client Id.
        /// </summary>
        public string ClientId { get; set; }

        /// <summary>
        /// Gets or sets Client secret.
        /// </summary>
        public string ClientSecret { get; set; }

        /// <summary>
        /// Gets or sets tenant id.
        /// </summary>
        public string TenantId { get; set; }

        /// <summary>
        /// Gets or sets Graph API scope.
        /// </summary>
        public string GraphScope { get; set; }

        /// <summary>
        /// Gets or sets Application Id URI.
        /// </summary>
        public string ApplicationIdUri { get; set; }

        /// <summary>
        /// Gets or sets valid issuers.
        /// </summary>
        public string ValidIssuers { get; set; }
    }
}
