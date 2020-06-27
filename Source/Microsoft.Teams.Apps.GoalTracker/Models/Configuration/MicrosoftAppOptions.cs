// <copyright file="MicrosoftAppOptions.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Models
{
    /// <summary>
    /// Provides configuration options for Microsoft Azure application registration
    /// </summary>
    public class MicrosoftAppOptions
    {
        /// <summary>
        /// Gets or Sets Azure AD application registration client id
        /// </summary>
        public string ClientId { get; set; }

        /// <summary>
        /// Gets or Sets Azure AD application registration client secret
        /// </summary>
        public string ClientSecret { get; set; }

        /// <summary>
        /// Gets or Sets Azure AD application registration tenant id
        /// </summary>
        public string TenantId { get; set; }
    }
}
