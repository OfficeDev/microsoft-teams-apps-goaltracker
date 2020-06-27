﻿// <copyright file="ITeamStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Common
{
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.GoalTracker.Models;

    /// <summary>
    /// Interface for team storage provider.
    /// </summary>
    public interface ITeamStorageProvider
    {
        /// <summary>
        /// Store or update team details in table storage.
        /// </summary>
        /// <param name="teamEntity">Represents team entity used for storage and retrieval.</param>
        /// <returns><see cref="Task"/> Returns the status whether team entity is stored or not.</returns>
        Task<bool> StoreOrUpdateTeamDetailAsync(TeamDetail teamEntity);

        /// <summary>
        /// Get already saved team entity from storage table.
        /// </summary>
        /// <param name="teamId">Team Id.</param>
        /// <returns><see cref="Task"/>Returns team entity.</returns>
        Task<TeamDetail> GetTeamDetailAsync(string teamId);

        /// <summary>
        /// This method delete the team detail record from table.
        /// </summary>
        /// <param name="teamEntity">Team configuration table entity.</param>
        /// <returns>A <see cref="Task"/> of type bool where true represents entity record is successfully deleted from table while false indicates failure in deleting data.</returns>
        Task<bool> DeleteTeamDetailAsync(TeamDetail teamEntity);
    }
}
