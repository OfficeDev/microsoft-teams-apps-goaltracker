// <copyright file="IPersonalGoalStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Common
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.GoalTracker.Models;

    /// <summary>
    /// Interface for provider class which helps in storing, updating, deleting personal goal details in storage.
    /// </summary>
    public interface IPersonalGoalStorageProvider
    {
        /// <summary>
        /// Stores or update personal goal detail in storage.
        /// </summary>
        /// <param name="personalGoalEntity">Holds personal goal data.</param>
        /// <returns>A boolean that represents personal goal detail entity data is saved or updated.</returns>
        Task<bool> CreateOrUpdatePersonalGoalDetailAsync(PersonalGoalDetail personalGoalEntity);

        /// <summary>
        /// Stores or update collection of personal goal details in storage.
        /// </summary>
        /// <param name="personalGoalEntities">Holds collection of personal goals data.</param>
        /// <returns>A boolean that represents personal goal detail entity data is saved or updated.</returns>
        Task<bool> CreateOrUpdatePersonalGoalDetailsAsync(IEnumerable<PersonalGoalDetail> personalGoalEntities);

        /// <summary>
        /// Get all personal goal details by AAD object id of user.
        /// </summary>
        /// <param name="userAadObjectId">AAD object id of the user for which personal goal details need to be fetched.</param>
        /// <returns>Returns collection of personal goal details.</returns>
        Task<IEnumerable<PersonalGoalDetail>> GetPersonalGoalDetailsByUserAadObjectIdAsync(string userAadObjectId);

        /// <summary>
        /// Get specific user's aligned goal details from personal goal detail storage table.
        /// </summary>
        /// <param name="teamId">Team id for which aligned goal details need to be fetched.</param>
        /// <param name="userAadObjectId">AAD object id of the user for which personal goal details need to be fetched.</param>
        /// <returns>Returns collection of aligned goal details for specific user.</returns>
        Task<IEnumerable<PersonalGoalDetail>> GetUserAlignedGoalDetailsByTeamIdAsync(string teamId, string userAadObjectId);

        /// <summary>
        /// Get personal goal reminder details which are unaligned for sending goal reminder.
        /// </summary>
        /// <returns>Returns collection of personal goal reminder details.</returns>
        Task<IEnumerable<PersonalGoalDetail>> GetPersonalUnalignedGoalReminderDetailsAsync();

        /// <summary>
        /// Get all personal goal details where IsDeleted flag is true.
        /// </summary>
        /// <returns>Returns collection of personal goal details where IsDeleted flag is true.</returns>
        Task<IEnumerable<PersonalGoalDetail>> GetPersonalDeletedGoalDetailsAsync();

        /// <summary>
        /// Delete personal goal details from storage.
        /// </summary>
        /// <param name="personalGoalEntities">Holds collection of personal goal details to be deleted from storage.</param>
        /// <returns>A boolean that represents personal goal details are deleted or not.</returns>
        Task<bool> DeletePersonalGoalDetailsAsync(IEnumerable<PersonalGoalDetail> personalGoalEntities);

        /// <summary>
        /// Get personal goal details by unique goal id.
        /// </summary>
        /// <param name="personalGoalId">Unique id of a personal goal.</param>
        /// <param name="userAadObjectId">AAD object id of the user for which personal goal details need to be fetched.</param>
        /// <returns>Returns personal goal detail for particular goal id.</returns>
        Task<PersonalGoalDetail> GetPersonalGoalDetailByGoalIdAsync(string personalGoalId, string userAadObjectId);
    }
}
