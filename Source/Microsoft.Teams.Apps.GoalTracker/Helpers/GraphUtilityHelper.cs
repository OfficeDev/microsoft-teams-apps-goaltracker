// <copyright file="GraphUtilityHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Helpers
{
    using System.Collections.Generic;
    using System.Net.Http.Headers;
    using System.Threading.Tasks;
    using Microsoft.Graph;
    using Microsoft.Teams.Apps.GoalTracker.Models;

    /// <summary>
    /// Implements the methods that are defined in <see cref="GraphUtilityHelper"/>.
    /// </summary>
    public class GraphUtilityHelper
    {
        /// <summary>
        /// Instance of graphServiceClient.
        /// </summary>
        private readonly GraphServiceClient graphClient;

        /// <summary>
        /// Initializes a new instance of the <see cref="GraphUtilityHelper"/> class.
        /// </summary>
        /// <param name="accessToken">User access token with Graph scopes.</param>
        public GraphUtilityHelper(
            string accessToken)
        {
            this.graphClient = new GraphServiceClient(
            new DelegateAuthenticationProvider(
            requestMessage =>
            {
                requestMessage.Headers.Authorization = new AuthenticationHeaderValue(
                    "Bearer",
                    accessToken);
                return Task.CompletedTask;
            }));
        }

        /// <summary>
        /// Get team owner details.
        /// </summary>
        /// <param name="groupId">Group id of the team of which owner needs to be listed.</param>
        /// <returns>A task that returns list of all channels in a team.</returns>
        public async Task<IEnumerable<TeamOwnerDetail>> GetTeamOwnerDetailsAsync(string groupId)
        {
            var teamOwnerDetails = new List<TeamOwnerDetail>();
            var owners = await this.graphClient.Groups[groupId].Owners.Request().GetAsync();
            foreach (var owner in owners)
            {
                // Add all team owner Ids to List.
                teamOwnerDetails.Add(new TeamOwnerDetail { TeamOwnerId = owner.Id });
            }

            return teamOwnerDetails;
        }
    }
}
