﻿// <copyright file="ITeamsInfoHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Helpers
{
    using System.Threading.Tasks;
    using Microsoft.Bot.Schema.Teams;

    /// <summary>
    /// Interface to provide team information helper methods.
    /// </summary>
    public interface ITeamsInfoHelper
    {
        /// <summary>
        /// To fetch team member information for specified team.
        /// </summary>
        /// <param name="teamId">Team id.</param>
        /// <param name="userId">User object id.</param>
        /// <returns>Team channel information.</returns>
        Task<TeamsChannelAccount> GetTeamMemberAsync(string teamId, string userId);
    }
}
