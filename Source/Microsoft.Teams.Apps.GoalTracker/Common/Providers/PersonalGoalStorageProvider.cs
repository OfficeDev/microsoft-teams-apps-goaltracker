// <copyright file="PersonalGoalStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Common
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Net;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.GoalTracker.Models;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Implements storage provider which helps in storing, updating, deleting personal goal details in storage.
    /// </summary>
    public class PersonalGoalStorageProvider : BaseStorageProvider, IPersonalGoalStorageProvider
    {
        /// <summary>
        /// Represents personal goal detail table name in storage.
        /// </summary>
        private const string PersonalGoalDetailTableName = "PersonalGoalDetail";

        /// <summary>
        /// Max number of goals for a batch operation.
        /// </summary>
        private const int PersonalGoalsPerBatch = 100;

        /// <summary>
        /// Instance to send logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<PersonalGoalStorageProvider> logger;

        /// <summary>
        /// Instance of search service for working with personal goal data in storage.
        /// </summary>
        private readonly IPersonalGoalSearchService personalGoalSearchService;

        /// <summary>
        /// Initializes a new instance of the <see cref="PersonalGoalStorageProvider"/> class.
        /// Handles storage read write operations.
        /// </summary>
        /// <param name="storageOptions">A set of key/value application configuration properties for storage.</param>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="personalGoalSearchService">Personal goal search service which will help in retrieving aligned goals information.</param>
        public PersonalGoalStorageProvider(
            IOptionsMonitor<StorageOptions> storageOptions,
            ILogger<PersonalGoalStorageProvider> logger,
            IPersonalGoalSearchService personalGoalSearchService)
            : base(storageOptions, PersonalGoalDetailTableName)
        {
            personalGoalSearchService = personalGoalSearchService ?? throw new ArgumentNullException(nameof(personalGoalSearchService));
            this.logger = logger;
            this.personalGoalSearchService = personalGoalSearchService;

            // Index creation takes time due to which search service is not able to fetch team goal status for the first time when application is deployed.
            // This will ensure index is created when application is deployed.
            this.personalGoalSearchService.EnsureInitializedAsync();
        }

        /// <summary>
        /// Stores or update personal goal detail in storage.
        /// </summary>
        /// <param name="personalGoalEntity">Holds personal goal data.</param>
        /// <returns>A boolean that represents personal goal detail entity data is saved or updated.</returns>
        public async Task<bool> CreateOrUpdatePersonalGoalDetailAsync(PersonalGoalDetail personalGoalEntity)
        {
            personalGoalEntity = personalGoalEntity ?? throw new ArgumentNullException(nameof(personalGoalEntity));

            try
            {
                await this.EnsureInitializedAsync();
                TableOperation operation = TableOperation.InsertOrReplace(personalGoalEntity);
                var result = await this.CloudTable.ExecuteAsync(operation);
                return result.HttpStatusCode == (int)HttpStatusCode.NoContent;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"An error occurred in {nameof(this.CreateOrUpdatePersonalGoalDetailAsync)} while saving personal goal data.");
                throw;
            }
        }

        /// <summary>
        /// Stores or update collection of personal goal details in storage.
        /// </summary>
        /// <param name="personalGoalEntities">Holds collection of personal goals data.</param>
        /// <returns>A boolean that represents personal goal detail entity data is saved or updated.</returns>
        public async Task<bool> CreateOrUpdatePersonalGoalDetailsAsync(IEnumerable<PersonalGoalDetail> personalGoalEntities)
        {
            personalGoalEntities = personalGoalEntities ?? throw new ArgumentNullException(nameof(personalGoalEntities));

            try
            {
                await this.EnsureInitializedAsync();
                TableBatchOperation tableBatchOperation = new TableBatchOperation();
                int batchCount = (int)Math.Ceiling((double)personalGoalEntities.Count() / PersonalGoalsPerBatch);
                for (int batchCountIndex = 0; batchCountIndex < batchCount; batchCountIndex++)
                {
                    var personalGoalDetailsBatch = personalGoalEntities.Skip(batchCountIndex * PersonalGoalsPerBatch).Take(PersonalGoalsPerBatch);
                    foreach (var personalGoalDetail in personalGoalDetailsBatch)
                    {
                        tableBatchOperation.InsertOrReplace(personalGoalDetail);
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
                this.logger.LogError(ex, $"An error occurred in {nameof(this.CreateOrUpdatePersonalGoalDetailsAsync)} while saving personal goal details in storage.");
                throw;
            }
        }

        /// <summary>
        /// Get all personal goal details by AAD object id of user.
        /// </summary>
        /// <param name="userAadObjectId">AAD object id of the user for which personal goal details need to be fetched.</param>
        /// <returns>Returns collection of personal goal details.</returns>
        public async Task<IEnumerable<PersonalGoalDetail>> GetPersonalGoalDetailsByUserAadObjectIdAsync(string userAadObjectId)
        {
            userAadObjectId = userAadObjectId ?? throw new ArgumentNullException(nameof(userAadObjectId));

            try
            {
                await this.EnsureInitializedAsync();
                string isActiveFilter = TableQuery.GenerateFilterConditionForBool(nameof(PersonalGoalDetail.IsActive), QueryComparisons.Equal, true);
                string isDeletedFilter = TableQuery.GenerateFilterConditionForBool(nameof(PersonalGoalDetail.IsDeleted), QueryComparisons.Equal, false);
                string userAadObjectIdFilter = TableQuery.GenerateFilterCondition(nameof(PersonalGoalDetail.PartitionKey), QueryComparisons.Equal, userAadObjectId);
                var query = new TableQuery<PersonalGoalDetail>().Where($"{isActiveFilter} and {isDeletedFilter} and {userAadObjectIdFilter}");
                return await this.GetPersonalGoalDetailsAsync(query);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"An error occurred in {nameof(this.GetPersonalGoalDetailsByUserAadObjectIdAsync)} while getting personal goal detail by AAD object id: {userAadObjectId}");
                throw;
            }
        }

        /// <summary>
        /// Get personal goal details by unique goal id.
        /// </summary>
        /// <param name="personalGoalId">Unique id of a personal goal.</param>
        /// <param name="userAadObjectId">AAD object id of the user for which personal goal details need to be fetched.</param>
        /// <returns>Returns personal goal detail for particular goal id.</returns>
        public async Task<PersonalGoalDetail> GetPersonalGoalDetailByGoalIdAsync(string personalGoalId, string userAadObjectId)
        {
            personalGoalId = personalGoalId ?? throw new ArgumentNullException(nameof(personalGoalId));
            userAadObjectId = userAadObjectId ?? throw new ArgumentNullException(nameof(userAadObjectId));

            try
            {
                await this.EnsureInitializedAsync();
                string isActiveFilter = TableQuery.GenerateFilterConditionForBool(nameof(PersonalGoalDetail.IsActive), QueryComparisons.Equal, true);
                string isDeletedFilter = TableQuery.GenerateFilterConditionForBool(nameof(PersonalGoalDetail.IsDeleted), QueryComparisons.Equal, false);
                string goalIdFilter = TableQuery.GenerateFilterCondition(nameof(PersonalGoalDetail.RowKey), QueryComparisons.Equal, personalGoalId);
                string userAadObjectIdFilter = TableQuery.GenerateFilterCondition(nameof(PersonalGoalDetail.PartitionKey), QueryComparisons.Equal, userAadObjectId);
                var query = new TableQuery<PersonalGoalDetail>().Where($"{goalIdFilter} and {userAadObjectIdFilter} and {isActiveFilter} and {isDeletedFilter}");
                var searchResult = await this.CloudTable.ExecuteQuerySegmentedAsync(query, null);
                return searchResult.Results.FirstOrDefault();
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"An error occurred in {nameof(this.GetPersonalGoalDetailByGoalIdAsync)} while getting personal goal detail by goal id: {personalGoalId}");
                throw;
            }
        }

        /// <summary>
        /// Get specific user's aligned goal details from personal goal detail storage table.
        /// </summary>
        /// <param name="teamId">Team id for which aligned goal details need to be fetched.</param>
        /// <param name="userAadObjectId">AAD object id of the user for which personal goal details need to be fetched.</param>
        /// <returns>Returns collection of aligned goal details for specific user.</returns>
        public async Task<IEnumerable<PersonalGoalDetail>> GetUserAlignedGoalDetailsByTeamIdAsync(string teamId, string userAadObjectId)
        {
            teamId = teamId ?? throw new ArgumentNullException(nameof(teamId));
            userAadObjectId = userAadObjectId ?? throw new ArgumentNullException(nameof(userAadObjectId));

            try
            {
                await this.EnsureInitializedAsync();
                string isActiveFilter = TableQuery.GenerateFilterConditionForBool(nameof(PersonalGoalDetail.IsActive), QueryComparisons.Equal, true);
                string isDeletedFilter = TableQuery.GenerateFilterConditionForBool(nameof(PersonalGoalDetail.IsDeleted), QueryComparisons.Equal, false);
                string isAlignFilter = TableQuery.GenerateFilterConditionForBool(nameof(PersonalGoalDetail.IsAligned), QueryComparisons.Equal, true);
                string teamIdFilter = TableQuery.GenerateFilterCondition(nameof(PersonalGoalDetail.TeamId), QueryComparisons.Equal, teamId);
                string userAadObjectIdFilter = TableQuery.GenerateFilterCondition(nameof(PersonalGoalDetail.PartitionKey), QueryComparisons.Equal, userAadObjectId);
                var query = new TableQuery<PersonalGoalDetail>().Where($"{isActiveFilter} and {isDeletedFilter} and {isAlignFilter} and {teamIdFilter} and {userAadObjectIdFilter}");
                return await this.GetPersonalGoalDetailsAsync(query);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"An error occurred in {nameof(this.GetUserAlignedGoalDetailsByTeamIdAsync)} for fetching user aligned goals for user: {userAadObjectId} and team : {teamId}");
                throw;
            }
        }

        /// <summary>
        /// Get personal goal reminder details which are unaligned for sending goal reminder.
        /// </summary>
        /// <returns>Returns collection of personal goal reminder details.</returns>
        public async Task<IEnumerable<PersonalGoalDetail>> GetPersonalUnalignedGoalReminderDetailsAsync()
        {
            try
            {
                await this.EnsureInitializedAsync();
                var today = DateTime.UtcNow;
                int weeklyReminder = today.DayOfWeek == CultureInfo.InvariantCulture.DateTimeFormat.FirstDayOfWeek + 1 ? 0 : -1; // Check if its Monday for weekly reminder frequency.
                int biweeklyReminder = today.Day == 1 || today.Day == 16 ? 1 : -1; // Check if its 1st or 16th day of month for bi-weekly reminder frequency.
                int monthlyReminder = today.Day == 1 ? 2 : -1; // Check if its 1st day of month for monthly reminder frequency.
                int quarterlyReminder = today.Day == 1 && ((today.Month + 2) % 3) == 0 ? 3 : -1; // Check if its 1st day of quarter for quarterly reminder frequency.

                // Get personal goals from storage whose end date or end date after 3 days is equal to today's date.
                string endDatePassedFilter = TableQuery.GenerateFilterCondition(nameof(PersonalGoalDetail.EndDateUTC), QueryComparisons.Equal, today.AddDays(-1).ToString(Constants.UTCDateFormat, CultureInfo.InvariantCulture));
                string endDateAfterThreeDaysFilter = TableQuery.GenerateFilterCondition(nameof(PersonalGoalDetail.EndDateUTC), QueryComparisons.Equal, today.AddDays(+3).ToString(Constants.UTCDateFormat, CultureInfo.InvariantCulture));

                // Get personal goals as per reminder frequency set.
                string weeklyReminderFilter = TableQuery.GenerateFilterConditionForInt(nameof(PersonalGoalDetail.ReminderFrequency), QueryComparisons.Equal, weeklyReminder);
                string biweeklyReminderFilter = TableQuery.GenerateFilterConditionForInt(nameof(PersonalGoalDetail.ReminderFrequency), QueryComparisons.Equal, biweeklyReminder);
                string monthlyReminderFilter = TableQuery.GenerateFilterConditionForInt(nameof(PersonalGoalDetail.ReminderFrequency), QueryComparisons.Equal, monthlyReminder);
                string quarterlyReminderFilter = TableQuery.GenerateFilterConditionForInt(nameof(PersonalGoalDetail.ReminderFrequency), QueryComparisons.Equal, quarterlyReminder);

                // Get only active and unaligned goals whose reminder frequency is active.
                string isActiveFilter = TableQuery.GenerateFilterConditionForBool(nameof(PersonalGoalDetail.IsActive), QueryComparisons.Equal, true);
                string isAlignedFilter = TableQuery.GenerateFilterConditionForBool(nameof(PersonalGoalDetail.IsAligned), QueryComparisons.Equal, false);
                string isDeletedFilter = TableQuery.GenerateFilterConditionForBool(nameof(PersonalGoalDetail.IsDeleted), QueryComparisons.Equal, false);
                string isReminderActiveFilter = TableQuery.GenerateFilterConditionForBool(nameof(PersonalGoalDetail.IsReminderActive), QueryComparisons.Equal, true);
                var query = new TableQuery<PersonalGoalDetail>().Where($"({endDatePassedFilter} or {endDateAfterThreeDaysFilter} or {weeklyReminderFilter} or {biweeklyReminderFilter} or {monthlyReminderFilter} or {quarterlyReminderFilter}) and {isAlignedFilter} and {isActiveFilter} and {isDeletedFilter} and {isReminderActiveFilter}");
                return await this.GetPersonalGoalDetailsAsync(query);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"An error occurred in {nameof(this.GetPersonalUnalignedGoalReminderDetailsAsync)}");
                throw;
            }
        }

        /// <summary>
        /// Get all personal goal details where IsDeleted flag is true.
        /// </summary>
        /// <returns>Returns collection of personal goal details where IsDeleted flag is true.</returns>
        public async Task<IEnumerable<PersonalGoalDetail>> GetPersonalDeletedGoalDetailsAsync()
        {
            try
            {
                await this.EnsureInitializedAsync();
                string isDeletedFilter = TableQuery.GenerateFilterConditionForBool(nameof(PersonalGoalDetail.IsDeleted), QueryComparisons.Equal, true);
                var query = new TableQuery<PersonalGoalDetail>().Where(isDeletedFilter);
                return await this.GetPersonalGoalDetailsAsync(query);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"An error occurred in {nameof(this.GetPersonalDeletedGoalDetailsAsync)}");
                throw;
            }
        }

        /// <summary>
        /// Delete personal goal details from storage.
        /// </summary>
        /// <param name="personalGoalEntities">Holds collection of personal goal details to be deleted from storage.</param>
        /// <returns>A boolean that represents personal goal details are deleted or not.</returns>
        public async Task<bool> DeletePersonalGoalDetailsAsync(IEnumerable<PersonalGoalDetail> personalGoalEntities)
        {
            personalGoalEntities = personalGoalEntities ?? throw new ArgumentNullException(nameof(personalGoalEntities));

            try
            {
                await this.EnsureInitializedAsync();
                TableBatchOperation tableBatchOperation = new TableBatchOperation();
                int batchCount = (int)Math.Ceiling((double)personalGoalEntities.Count() / PersonalGoalsPerBatch);
                for (int batchCountIndex = 0; batchCountIndex < batchCount; batchCountIndex++)
                {
                    var personalGoalEntitiesBatch = personalGoalEntities.Skip(batchCountIndex * PersonalGoalsPerBatch).Take(PersonalGoalsPerBatch);
                    foreach (var personalGoalEntity in personalGoalEntitiesBatch)
                    {
                        tableBatchOperation.Delete(personalGoalEntity);
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
                this.logger.LogError(ex, $"An error occurred in {nameof(this.DeletePersonalGoalDetailsAsync)} while deleting personal goal details.");
                throw;
            }
        }

        /// <summary>
        /// Get personal goal details from storage depending upon filter condition.
        /// </summary>
        /// <param name="query">Query condition to fetch data from storage.</param>
        /// <returns>Returns collection of personal goal details from storage.</returns>
        private async Task<IEnumerable<PersonalGoalDetail>> GetPersonalGoalDetailsAsync(TableQuery<PersonalGoalDetail> query)
        {
            TableContinuationToken continuationToken = null;
            var personalGoalDetails = new List<PersonalGoalDetail>();

            do
            {
                var queryResult = await this.CloudTable.ExecuteQuerySegmentedAsync(query, continuationToken);
                if (queryResult?.Results != null)
                {
                    personalGoalDetails.AddRange(queryResult.Results);
                    continuationToken = queryResult.ContinuationToken;
                }
            }
            while (continuationToken != null);

            return personalGoalDetails;
        }
    }
}
