// <copyright file="PersonalGoalNoteStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Common
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.GoalTracker.Models;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Implements storage provider which helps in storing, updating, deleting personal note detail in storage.
    /// </summary>
    public class PersonalGoalNoteStorageProvider : BaseStorageProvider, IPersonalGoalNoteStorageProvider
    {
        /// <summary>
        /// Represents personal note detail table name in storage.
        /// </summary>
        private const string PersonalGoalNoteDetailTableName = "PersonalGoalNoteDetail";

        /// <summary>
        /// Max number of notes for a batch operation.
        /// </summary>
        private const int PersonalGoalNotesPerBatch = 100;

        /// <summary>
        /// Instance to send logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<PersonalGoalNoteStorageProvider> logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="PersonalGoalNoteStorageProvider"/> class.
        /// Handles storage read write operations.
        /// </summary>
        /// <param name="storageOptions">A set of key/value application configuration properties for storage.</param>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        public PersonalGoalNoteStorageProvider(IOptionsMonitor<StorageOptions> storageOptions, ILogger<PersonalGoalNoteStorageProvider> logger)
            : base(storageOptions, PersonalGoalNoteDetailTableName)
        {
            this.logger = logger;
        }

        /// <summary>
        /// Stores or update collection of notes for individual personal goal in storage.
        /// </summary>
        /// <param name="personalGoalNoteEntities">Holds collection of personal goal notes data.</param>
        /// <returns>A boolean that represents personal goal note detail entities are saved or updated.</returns>
        public async Task<bool> CreateOrUpdatePersonalGoalNoteDetailsAsync(IEnumerable<PersonalGoalNoteDetail> personalGoalNoteEntities)
        {
            personalGoalNoteEntities = personalGoalNoteEntities ?? throw new ArgumentNullException(nameof(personalGoalNoteEntities));

            try
            {
                await this.EnsureInitializedAsync();
                TableBatchOperation tableBatchOperation = new TableBatchOperation();
                int batchCount = (int)Math.Ceiling((double)personalGoalNoteEntities.Count() / PersonalGoalNotesPerBatch);
                for (int batchCountIndex = 0; batchCountIndex < batchCount; batchCountIndex++)
                {
                    var personalGoalNoteEntitiesBatch = personalGoalNoteEntities.Skip(batchCountIndex * PersonalGoalNotesPerBatch).Take(PersonalGoalNotesPerBatch);
                    foreach (var personalGoalNoteEntity in personalGoalNoteEntitiesBatch)
                    {
                        tableBatchOperation.InsertOrReplace(personalGoalNoteEntity);
                    }

                    if (tableBatchOperation.Count > 0)
                    {
                        await this.CloudTable.ExecuteBatchAsync(tableBatchOperation);
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"An error occurred in {nameof(this.CreateOrUpdatePersonalGoalNoteDetailsAsync)} while storing notes in storage.");
                throw;
            }
        }

        /// <summary>
        /// Stores or updates single note for personal goal in storage.
        /// </summary>
        /// <param name="personalGoalNoteDetail">Holds personal goal note data.</param>
        /// <returns>A task that represents personal goal note detail entity data is saved or updated.</returns>
        public async Task<bool> CreateOrUpdatePersonalGoalNoteDetailAsync(PersonalGoalNoteDetail personalGoalNoteDetail)
        {
            personalGoalNoteDetail = personalGoalNoteDetail ?? throw new ArgumentNullException(nameof(personalGoalNoteDetail));

            try
            {
                await this.EnsureInitializedAsync();
                TableOperation operation = TableOperation.InsertOrReplace(personalGoalNoteDetail);
                var result = await this.CloudTable.ExecuteAsync(operation);
                return true;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"An error occurred in {nameof(this.CreateOrUpdatePersonalGoalNoteDetailAsync)} while storing a note in storage for personal goal id: {personalGoalNoteDetail.PersonalGoalId}.");
                throw;
            }
        }

        /// <summary>
        /// Update collection of notes for individual personal goal in storage.
        /// </summary>
        /// <param name="personalGoalNoteEntities">Holds collection of personal goal notes data.</param>
        /// <returns>A boolean that represents personal goal note detail entities are updated.</returns>
        public async Task<bool> UpdatePersonalGoalNoteDetailsAsync(IEnumerable<PersonalGoalNoteDetail> personalGoalNoteEntities)
        {
            personalGoalNoteEntities = personalGoalNoteEntities ?? throw new ArgumentNullException(nameof(personalGoalNoteEntities));

            try
            {
                await this.EnsureInitializedAsync();
                TableBatchOperation tableBatchOperation = new TableBatchOperation();
                int batchCount = (int)Math.Ceiling((double)personalGoalNoteEntities.Count() / PersonalGoalNotesPerBatch);
                for (int batchCountIndex = 0; batchCountIndex < batchCount; batchCountIndex++)
                {
                    var personalGoalNoteEntitiesBatch = personalGoalNoteEntities.Skip(batchCountIndex * PersonalGoalNotesPerBatch).Take(PersonalGoalNotesPerBatch);
                    foreach (var personalGoalNoteEntity in personalGoalNoteEntitiesBatch)
                    {
                        tableBatchOperation.Replace(personalGoalNoteEntity);
                    }

                    if (tableBatchOperation.Count > 0)
                    {
                        await this.CloudTable.ExecuteBatchAsync(tableBatchOperation);
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"An error occurred in {nameof(this.CreateOrUpdatePersonalGoalNoteDetailsAsync)} while storing notes in storage.");
                throw;
            }
        }

        /// <summary>
        /// Get already saved personal goal note detail from storage table.
        /// </summary>
        /// <param name="personalGoalNoteId">Personal goal note id based on which note data will be fetched.</param>
        /// <param name="userAadObjectId">AAD object id of the user for which personal goal note details need to be fetched.</param>
        /// <returns><see cref="Task"/>Already saved goal note entity details.</returns>
        public async Task<PersonalGoalNoteDetail> GetPersonalGoalNoteDetailAsync(string personalGoalNoteId, string userAadObjectId)
        {
            personalGoalNoteId = personalGoalNoteId ?? throw new ArgumentNullException(nameof(personalGoalNoteId));
            userAadObjectId = userAadObjectId ?? throw new ArgumentNullException(nameof(userAadObjectId));

            try
            {
                await this.EnsureInitializedAsync();
                string isActiveFilter = TableQuery.GenerateFilterConditionForBool(nameof(PersonalGoalNoteDetail.IsActive), QueryComparisons.Equal, true);
                string userAadObjectIdFilter = TableQuery.GenerateFilterCondition(nameof(PersonalGoalNoteDetail.PartitionKey), QueryComparisons.Equal, userAadObjectId);
                string goalNoteIdFilter = TableQuery.GenerateFilterCondition(nameof(PersonalGoalNoteDetail.RowKey), QueryComparisons.Equal, personalGoalNoteId);
                var query = new TableQuery<PersonalGoalNoteDetail>().Where($"{userAadObjectIdFilter} and {isActiveFilter} and {goalNoteIdFilter}");
                var searchResult = await this.CloudTable.ExecuteQuerySegmentedAsync(query, null);
                return searchResult.Results.FirstOrDefault();
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"An error occurred in {nameof(this.GetPersonalGoalNoteDetailAsync)} while getting personal note detail for personalGoalNoteId: {personalGoalNoteId}");
                throw;
            }
        }

        /// <summary>
        /// Get number of notes added for a personal goal.
        /// </summary>
        /// <param name="personalGoalId">Unique goal id for each personal goal.</param>
        /// <param name="userAadObjectId">AAD object id of the user for which personal goal note details need to be fetched.</param>
        /// <returns><see cref="Task"/> Number of notes added for particular goal.</returns>
        public async Task<int> GetNumberOfNotesForGoalAsync(string personalGoalId, string userAadObjectId)
        {
            personalGoalId = personalGoalId ?? throw new ArgumentNullException(nameof(personalGoalId));
            userAadObjectId = userAadObjectId ?? throw new ArgumentNullException(nameof(userAadObjectId));

            try
            {
                await this.EnsureInitializedAsync();
                string isActiveFilter = TableQuery.GenerateFilterConditionForBool(nameof(PersonalGoalDetail.IsActive), QueryComparisons.Equal, true);
                string goalIdFilter = TableQuery.GenerateFilterCondition(nameof(PersonalGoalNoteDetail.PersonalGoalId), QueryComparisons.Equal, personalGoalId);
                string userAadObjectIdFilter = TableQuery.GenerateFilterCondition(nameof(PersonalGoalNoteDetail.PartitionKey), QueryComparisons.Equal, userAadObjectId);
                var query = new TableQuery<PersonalGoalNoteDetail>().Where($"{userAadObjectIdFilter} and {goalIdFilter} and {isActiveFilter} ");
                var searchResult = await this.CloudTable.ExecuteQuerySegmentedAsync(query, null);
                return searchResult.Results.Count;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"An error occurred in {nameof(this.GetNumberOfNotesForGoalAsync)} while getting number of personal goal notes by personal goal Id: {personalGoalId}");
                throw;
            }
        }

        /// <summary>
        /// Get all personal goal note details added by user by user AAD object id.
        /// </summary>
        /// <param name="userAadObjectId">AAD object id of the user for which personal goal details need to be fetched.</param>
        /// <returns>Returns collection of personal goal note details.</returns>
        public async Task<IEnumerable<PersonalGoalNoteDetail>> GetPersonalGoalNoteDetailsByUserAadObjectIdAsync(string userAadObjectId)
        {
            userAadObjectId = userAadObjectId ?? throw new ArgumentNullException(nameof(userAadObjectId));

            try
            {
                await this.EnsureInitializedAsync();
                string isActiveFilter = TableQuery.GenerateFilterConditionForBool(nameof(PersonalGoalNoteDetail.IsActive), QueryComparisons.Equal, true);
                string userAadObjectIdFilter = TableQuery.GenerateFilterCondition(nameof(PersonalGoalNoteDetail.PartitionKey), QueryComparisons.Equal, userAadObjectId);
                var query = new TableQuery<PersonalGoalNoteDetail>().Where($"{isActiveFilter} and {userAadObjectIdFilter}");
                return await this.GetPersonalGoalNoteDetailsAsync(query);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"An error occurred in {nameof(this.GetPersonalGoalNoteDetailsByUserAadObjectIdAsync)} while getting personal goal detail by AAD object id: {userAadObjectId}");
                throw;
            }
        }

        /// <summary>
        /// Get personal goal note details for specific personal goal Id.
        /// </summary>
        /// <param name="personalGoalId">Unique id of a personal goal.</param>
        /// <param name="userAadObjectId">AAD object id of the user for which personal goal details need to be fetched.</param>
        /// <returns>Returns collection of personal goal note details.</returns>
        public async Task<IEnumerable<PersonalGoalNoteDetail>> GetPersonalGoalNoteDetailsAsync(string personalGoalId, string userAadObjectId)
        {
            personalGoalId = personalGoalId ?? throw new ArgumentNullException(nameof(personalGoalId));
            userAadObjectId = userAadObjectId ?? throw new ArgumentNullException(nameof(userAadObjectId));

            try
            {
                await this.EnsureInitializedAsync();
                string isActiveFilter = TableQuery.GenerateFilterConditionForBool(nameof(PersonalGoalNoteDetail.IsActive), QueryComparisons.Equal, true);
                string personalGoalIdFilter = TableQuery.GenerateFilterCondition(nameof(PersonalGoalNoteDetail.PersonalGoalId), QueryComparisons.Equal, personalGoalId);
                string userAadObjectIdFilter = TableQuery.GenerateFilterCondition(nameof(PersonalGoalNoteDetail.PartitionKey), QueryComparisons.Equal, userAadObjectId);
                var query = new TableQuery<PersonalGoalNoteDetail>().Where($"{personalGoalIdFilter} and {isActiveFilter}");
                return await this.GetPersonalGoalNoteDetailsAsync(query);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"An error occurred in {nameof(this.GetPersonalGoalNoteDetailsAsync)} while getting personal goal detail by personal goal id: {personalGoalId}");
                throw;
            }
        }

        /// <summary>
        /// Delete notes for personal goals from storage.
        /// </summary>
        /// <param name="personalGoalNoteEntities">Holds collection of personal goal note details.</param>
        /// <returns>A boolean that represents personal goal note detail entities are deleted or not.</returns>
        public async Task<bool> DeletePersonalGoalNoteDetailsAsync(IEnumerable<PersonalGoalNoteDetail> personalGoalNoteEntities)
        {
            personalGoalNoteEntities = personalGoalNoteEntities ?? throw new ArgumentNullException(nameof(personalGoalNoteEntities));

            try
            {
                await this.EnsureInitializedAsync();
                TableBatchOperation tableBatchOperation = new TableBatchOperation();
                int batchCount = (int)Math.Ceiling((double)personalGoalNoteEntities.Count() / PersonalGoalNotesPerBatch);
                for (int batchCountIndex = 0; batchCountIndex < batchCount; batchCountIndex++)
                {
                    var personalGoalNoteEntitiesBatch = personalGoalNoteEntities.Skip(batchCountIndex * PersonalGoalNotesPerBatch).Take(PersonalGoalNotesPerBatch);
                    foreach (var personalGoalNoteEntity in personalGoalNoteEntitiesBatch)
                    {
                        tableBatchOperation.Delete(personalGoalNoteEntity);
                    }

                    if (tableBatchOperation.Count > 0)
                    {
                        await this.CloudTable.ExecuteBatchAsync(tableBatchOperation);
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"An error occurred in {nameof(this.DeletePersonalGoalNoteDetailsAsync)}: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Get personal goal note details from storage depending upon query provided.
        /// </summary>
        /// <param name="query">Query condition to fetch data from storage.</param>
        /// <returns>Returns collection of personal goal note details from storage.</returns>
        private async Task<IEnumerable<PersonalGoalNoteDetail>> GetPersonalGoalNoteDetailsAsync(TableQuery<PersonalGoalNoteDetail> query)
        {
            TableContinuationToken continuationToken = null;
            var personalGoalNoteDetails = new List<PersonalGoalNoteDetail>();

            do
            {
                var queryResult = await this.CloudTable.ExecuteQuerySegmentedAsync(query, continuationToken);
                if (queryResult?.Results != null)
                {
                    personalGoalNoteDetails.AddRange(queryResult.Results);
                    continuationToken = queryResult.ContinuationToken;
                }
            }
            while (continuationToken != null);

            return personalGoalNoteDetails;
        }
    }
}
