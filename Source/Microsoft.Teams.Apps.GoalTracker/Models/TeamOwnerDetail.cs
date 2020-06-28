// <copyright file="TeamOwnerDetail.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Models
{
    using Newtonsoft.Json;

    /// <summary>
    /// Model for team owner details.
    /// </summary>
    public class TeamOwnerDetail
    {
        /// <summary>
        /// Gets or sets the team owner AAD object id.
        /// </summary>
        [JsonProperty("TeamOwnerId")]
        public string TeamOwnerId { get; set; }
    }
}
