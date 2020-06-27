// <copyright file="ITeamGoalStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Common
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.GoalTracker.Models;

    /// <summary>
    /// Interface for provider class which helps in storing, updating, deleting team goal details in storage.
    /// </summary>
    public interface ITeamGoalStorageProvider
    {
        /// <summary>
        /// Stores or updates team goal details in storage.
        /// </summary>
        /// <param name="teamGoalEntities">Holds collection of team goal details.</param>
        /// <returns>A boolean that represents team goal details are saved or updated.</returns>
        Task<bool> CreateOrUpdateTeamGoalDetailsAsync(IEnumerable<TeamGoalDetail> teamGoalEntities);

        /// <summary>
        /// Get team goal details by Microsoft Teams' team Id.
        /// </summary>
        /// <param name="teamId">Team id for which team goal details need to be fetched.</param>
        /// <returns>Returns collection of team goal details.</returns>
        Task<IEnumerable<TeamGoalDetail>> GetTeamGoalDetailsByTeamIdAsync(string teamId);

        /// <summary>
        /// Get specific team goal detail by unique team goal id.
        /// </summary>
        /// <param name="teamGoalId">Team goal id for which team goal details need to be fetched.</param>
        /// <param name="teamId">Team id for which team goal details need to be fetched.</param>
        /// <returns>Returns collection of team goal details.</returns>
        Task<TeamGoalDetail> GetTeamGoalDetailByTeamGoalIdAsync(string teamGoalId, string teamId);

        /// <summary>
        /// Get all team goal details for sending goal reminder.
        /// </summary>
        /// <returns>Returns collection of team goal details.</returns>
        Task<IEnumerable<TeamGoalDetail>> GetTeamGoalReminderDetailsAsync();

        /// <summary>
        /// Get all team goal details where IsDeleted flag is true.
        /// </summary>
        /// <returns>Returns collection of team goal details where IsDeleted flag is true.</returns>
        Task<IEnumerable<TeamGoalDetail>> GetDeletedTeamGoalDetailsAsync();

        /// <summary>
        /// Delete team goal details data from storage.
        /// </summary>
        /// <param name="teamGoalEntities">Holds collection of team goal details.</param>
        /// <returns>A boolean that represents team goal detail entity data is deleted.</returns>
        Task<bool> DeleteTeamGoalDetailsAsync(IEnumerable<TeamGoalDetail> teamGoalEntities);
    }
}
