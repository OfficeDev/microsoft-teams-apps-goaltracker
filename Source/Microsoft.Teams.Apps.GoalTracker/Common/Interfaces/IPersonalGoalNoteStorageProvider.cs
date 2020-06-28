// <copyright file="IPersonalGoalNoteStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Common
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.GoalTracker.Models;

    /// <summary>
    /// Interface for provider class which helps in storing, updating, deleting personal note detail in storage.
    /// </summary>
    public interface IPersonalGoalNoteStorageProvider
    {
        /// <summary>
        /// Stores or update collection of notes for individual personal goal in storage.
        /// </summary>
        /// <param name="personalGoalNoteEntities">Holds collection of personal goal notes data.</param>
        /// <returns>A boolean that represents personal goal note detail entities are saved or updated.</returns>
        Task<bool> CreateOrUpdatePersonalGoalNoteDetailsAsync(IEnumerable<PersonalGoalNoteDetail> personalGoalNoteEntities);

        /// <summary>
        /// Updates collection of notes for individual personal goal in storage.
        /// </summary>
        /// <param name="personalGoalNoteEntities">Holds collection of personal goal notes data.</param>
        /// <returns>A boolean that represents personal goal note detail entities are updated.</returns>
        Task<bool> UpdatePersonalGoalNoteDetailsAsync(IEnumerable<PersonalGoalNoteDetail> personalGoalNoteEntities);

        /// <summary>
        /// Delete notes for personal goals from storage.
        /// </summary>
        /// <param name="personalGoalNoteEntities">Holds collection of personal goal note details.</param>
        /// <returns>A boolean that represents personal goal note detail entities are deleted or not.</returns>
        Task<bool> DeletePersonalGoalNoteDetailsAsync(IEnumerable<PersonalGoalNoteDetail> personalGoalNoteEntities);

        /// <summary>
        /// Stores or updates single note for personal goal in storage.
        /// </summary>
        /// <param name="personalGoalNoteDetail">Holds personal goal note data.</param>
        /// <returns>A task that represents personal goal note detail entity data is saved or updated.</returns>
        Task<bool> CreateOrUpdatePersonalGoalNoteDetailAsync(PersonalGoalNoteDetail personalGoalNoteDetail);

        /// <summary>
        /// Get already saved personal goal note detail from storage table.
        /// </summary>
        /// <param name="personalGoalNoteId">Personal goal note id based on which note data will be fetched.</param>
        /// <param name="userAadObjectId">AAD object id of the user for which personal goal note details need to be fetched.</param>
        /// <returns><see cref="Task"/>Already saved goal note entity details.</returns>
        Task<PersonalGoalNoteDetail> GetPersonalGoalNoteDetailAsync(string personalGoalNoteId, string userAadObjectId);

        /// <summary>
        /// Get number of notes added for a personal goal.
        /// </summary>
        /// <param name="personalGoalId">Unique goal id for each personal goal.</param>
        /// <param name="userAadObjectId">AAD object id of the user for which personal goal note details need to be fetched.</param>
        /// <returns><see cref="Task"/> Number of notes added for particular goal.</returns>
        Task<int> GetNumberOfNotesForGoalAsync(string personalGoalId, string userAadObjectId);

        /// <summary>
        /// Get all personal goal note details added by user by user AAD object id.
        /// </summary>
        /// <param name="userAadObjectId">AAD object id of the user for which personal goal details need to be fetched.</param>
        /// <returns>Returns collection of personal goal note details.</returns>
        Task<IEnumerable<PersonalGoalNoteDetail>> GetPersonalGoalNoteDetailsByUserAadObjectIdAsync(string userAadObjectId);

        /// <summary>
        /// Get personal goal note details for specific personal goal Id.
        /// </summary>
        /// <param name="personalGoalId">Unique id of a personal goal.</param>
        /// <param name="userAadObjectId">AAD object id of the user for which personal goal details need to be fetched.</param>
        /// <returns>Returns collection of personal goal note details.</returns>
        Task<IEnumerable<PersonalGoalNoteDetail>> GetPersonalGoalNoteDetailsAsync(string personalGoalId, string userAadObjectId);
    }
}
