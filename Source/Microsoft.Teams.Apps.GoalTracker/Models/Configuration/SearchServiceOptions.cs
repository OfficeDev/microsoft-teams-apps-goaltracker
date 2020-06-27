// <copyright file="SearchServiceOptions.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Models
{
    /// <summary>
    /// Provides settings related to search service.
    /// </summary>
    public class SearchServiceOptions
    {
        /// <summary>
        /// Gets or sets search service name.
        /// </summary>
        public string SearchServiceName { get; set; }

        /// <summary>
        /// Gets or sets search service query API key.
        /// </summary>
        public string SearchServiceQueryApiKey { get; set; }

        /// <summary>
        /// Gets or sets search service admin API key.
        /// </summary>
        public string SearchServiceAdminApiKey { get; set; }
    }
}
