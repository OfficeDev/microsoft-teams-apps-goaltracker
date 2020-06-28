// <copyright file="TeamGoalStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Common
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.GoalTracker.Models;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Implements storage provider which helps in storing, updating, deleting team goal details in storage.
    /// </summary>
    public class TeamGoalStorageProvider : BaseStorageProvider, ITeamGoalStorageProvider
    {
        /// <summary>
        /// Represents team goal detail table name in storage.
        /// </summary>
        private const string TeamGoalDetailTableName = "TeamGoalDetail";

        /// <summary>
        /// Max number of team goals for a batch operation.
        /// </summary>
        private const int TeamGoalsPerBatch = 100;

        /// <summary>
        /// Instance to send logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<TeamGoalStorageProvider> logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="TeamGoalStorageProvider"/> class.
        /// Handles storage read write operations.
        /// </summary>
        /// <param name="storageOptions">A set of key/value application configuration properties for storage.</param>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        public TeamGoalStorageProvider(IOptionsMonitor<StorageOptions> storageOptions, ILogger<TeamGoalStorageProvider> logger)
            : base(storageOptions, TeamGoalDetailTableName)
        {
            this.logger = logger;
        }

        /// <summary>
        /// Stores or updates team goal details in storage.
        /// </summary>
        /// <param name="teamGoalEntities">Holds collection of team goal details.</param>
        /// <returns>A boolean that represents team goal details are saved or updated.</returns>
        public async Task<bool> CreateOrUpdateTeamGoalDetailsAsync(IEnumerable<TeamGoalDetail> teamGoalEntities)
        {
            teamGoalEntities = teamGoalEntities ?? throw new ArgumentNullException(nameof(teamGoalEntities));

            try
            {
                await this.EnsureInitializedAsync();
                TableBatchOperation tableBatchOperation = new TableBatchOperation();
                int batchCount = (int)Math.Ceiling((double)teamGoalEntities.Count() / TeamGoalsPerBatch);
                for (int batchCountIndex = 0; batchCountIndex < batchCount; batchCountIndex++)
                {
                    var teamGoalEntitiesBatch = teamGoalEntities.Skip(batchCountIndex * TeamGoalsPerBatch).Take(TeamGoalsPerBatch);
                    foreach (var teamGoalEntity in teamGoalEntitiesBatch)
                    {
                        tableBatchOperation.InsertOrReplace(teamGoalEntity);
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
                this.logger.LogError(ex, $"An error occurred in {nameof(this.CreateOrUpdateTeamGoalDetailsAsync)} while saving team goal details.");
                throw;
            }
        }

        /// <summary>
        /// Get team goal details by Microsoft Teams' team Id.
        /// </summary>
        /// <param name="teamId">Team id for which team goal details need to be fetched.</param>
        /// <returns>Returns collection of team goal details.</returns>
        public async Task<IEnumerable<TeamGoalDetail>> GetTeamGoalDetailsByTeamIdAsync(string teamId)
        {
            teamId = teamId ?? throw new ArgumentNullException(nameof(teamId));

            try
            {
                await this.EnsureInitializedAsync();
                string isActiveFilter = TableQuery.GenerateFilterConditionForBool(nameof(TeamGoalDetail.IsActive), QueryComparisons.Equal, true);
                string isDeletedFilter = TableQuery.GenerateFilterConditionForBool(nameof(TeamGoalDetail.IsDeleted), QueryComparisons.Equal, false);
                string teamIdFilter = TableQuery.GenerateFilterCondition(nameof(TeamGoalDetail.PartitionKey), QueryComparisons.Equal, teamId);
                var query = new TableQuery<TeamGoalDetail>().Where($"{isActiveFilter} and {isDeletedFilter} and {teamIdFilter}");
                return await this.GetTeamGoalDetailsAsync(query);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"An error occurred in {nameof(this.GetTeamGoalDetailsByTeamIdAsync)} while getting team goal details by team id: {teamId}");
                throw;
            }
        }

        /// <summary>
        /// Get specific team goal detail by unique team goal id.
        /// </summary>
        /// <param name="teamGoalId">Team goal id for which team goal details need to be fetched.</param>
        /// <param name="teamId">Team id for which team goal details need to be fetched.</param>
        /// <returns>Returns collection of team goal details.</returns>
        public async Task<TeamGoalDetail> GetTeamGoalDetailByTeamGoalIdAsync(string teamGoalId, string teamId)
        {
            teamGoalId = teamGoalId ?? throw new ArgumentNullException(nameof(teamGoalId));
            teamId = teamId ?? throw new ArgumentNullException(nameof(teamId));

            try
            {
                await this.EnsureInitializedAsync();
                string teamGoalIdFilter = TableQuery.GenerateFilterCondition(nameof(TeamGoalDetail.RowKey), QueryComparisons.Equal, teamGoalId);
                string teamIdFilter = TableQuery.GenerateFilterCondition(nameof(TeamGoalDetail.PartitionKey), QueryComparisons.Equal, teamId);
                var query = new TableQuery<TeamGoalDetail>().Where($"{teamGoalIdFilter} and {teamIdFilter}");
                var searchResult = await this.CloudTable.ExecuteQuerySegmentedAsync(query, null);
                return searchResult.Results.FirstOrDefault();
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"An error occurred in {nameof(this.GetTeamGoalDetailsByTeamIdAsync)} while getting team goal details by team id: {teamId}");
                throw;
            }
        }

        /// <summary>
        /// Get all team goal details for sending goal reminder.
        /// </summary>
        /// <returns>Returns collection of team goal details.</returns>
        public async Task<IEnumerable<TeamGoalDetail>> GetTeamGoalReminderDetailsAsync()
        {
            try
            {
                await this.EnsureInitializedAsync();
                var today = DateTime.UtcNow;
                int weeklyReminder = today.DayOfWeek == CultureInfo.CurrentCulture.DateTimeFormat.FirstDayOfWeek + 1 ? 0 : -1; // Check if its Monday for weekly reminder frequency.
                int biweeklyReminder = today.Day == 1 || today.Day == 16 ? 1 : -1; // Check if its 1st or 16th day of month for bi-weekly reminder frequency.
                int monthlyReminder = today.Day == 1 ? 2 : -1; // Check if its 1st day of month for monthly reminder frequency.
                int quarterlyReminder = today.Day == 1 && ((today.Month + 2) % 3) == 0 ? 3 : -1; // Check if its 1st day of quarter for quarterly reminder frequency.

                // Get team goals from storage whose end date or end date after 3 days is equal to today's date.
                string endDatePassedFilter = TableQuery.GenerateFilterCondition(nameof(TeamGoalDetail.TeamGoalEndDateUTC), QueryComparisons.Equal, today.AddDays(-1).ToString(Constants.UTCDateFormat, CultureInfo.InvariantCulture));
                string endDateAfterThreeDaysFilter = TableQuery.GenerateFilterCondition(nameof(TeamGoalDetail.TeamGoalEndDateUTC), QueryComparisons.Equal, today.AddDays(+3).ToString(Constants.UTCDateFormat, CultureInfo.InvariantCulture));

                // Get team goals as per reminder frequency set.
                string weeklyReminderFilter = TableQuery.GenerateFilterConditionForInt(nameof(TeamGoalDetail.ReminderFrequency), QueryComparisons.Equal, weeklyReminder);
                string biweeklyReminderFilter = TableQuery.GenerateFilterConditionForInt(nameof(TeamGoalDetail.ReminderFrequency), QueryComparisons.Equal, biweeklyReminder);
                string monthlyReminderFilter = TableQuery.GenerateFilterConditionForInt(nameof(TeamGoalDetail.ReminderFrequency), QueryComparisons.Equal, monthlyReminder);
                string quarterlyReminderFilter = TableQuery.GenerateFilterConditionForInt(nameof(TeamGoalDetail.ReminderFrequency), QueryComparisons.Equal, quarterlyReminder);

                // Get only active goals whose reminder frequency is active.
                string isActiveFilter = TableQuery.GenerateFilterConditionForBool(nameof(TeamGoalDetail.IsActive), QueryComparisons.Equal, true);
                string isDeletedFilter = TableQuery.GenerateFilterConditionForBool(nameof(TeamGoalDetail.IsDeleted), QueryComparisons.Equal, false);
                string isReminderActiveFilter = TableQuery.GenerateFilterConditionForBool(nameof(TeamGoalDetail.IsReminderActive), QueryComparisons.Equal, true);
                var query = new TableQuery<TeamGoalDetail>().Where($"({endDatePassedFilter} or {endDateAfterThreeDaysFilter} or {weeklyReminderFilter} or {biweeklyReminderFilter} or {monthlyReminderFilter} or {quarterlyReminderFilter}) and {isActiveFilter} and {isDeletedFilter} and {isReminderActiveFilter}");
                return await this.GetTeamGoalDetailsAsync(query);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"An error occurred in {nameof(this.GetTeamGoalReminderDetailsAsync)} while getting team goal details.");
                throw;
            }
        }

        /// <summary>
        /// Get all team goal details where IsDeleted flag is true.
        /// </summary>
        /// <returns>Returns collection of team goal details where IsDeleted flag is true.</returns>
        public async Task<IEnumerable<TeamGoalDetail>> GetDeletedTeamGoalDetailsAsync()
        {
            try
            {
                await this.EnsureInitializedAsync();
                string isDeletedFilter = TableQuery.GenerateFilterConditionForBool(nameof(TeamGoalDetail.IsDeleted), QueryComparisons.Equal, true);
                var query = new TableQuery<TeamGoalDetail>().Where(isDeletedFilter);
                return await this.GetTeamGoalDetailsAsync(query);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"An error occurred in {nameof(this.GetDeletedTeamGoalDetailsAsync)}");
                throw;
            }
        }

        /// <summary>
        /// Delete team goal details data from storage.
        /// </summary>
        /// <param name="teamGoalEntities">Holds collection of team goal details.</param>
        /// <returns>A boolean that represents team goal detail entity data is deleted.</returns>
        public async Task<bool> DeleteTeamGoalDetailsAsync(IEnumerable<TeamGoalDetail> teamGoalEntities)
        {
            teamGoalEntities = teamGoalEntities ?? throw new ArgumentNullException(nameof(teamGoalEntities));

            try
            {
                await this.EnsureInitializedAsync();
                TableBatchOperation tableBatchOperation = new TableBatchOperation();
                int batchCount = (int)Math.Ceiling((double)teamGoalEntities.Count() / TeamGoalsPerBatch);
                for (int batchCountIndex = 0; batchCountIndex < batchCount; batchCountIndex++)
                {
                    var teamGoalEntitiesBatch = teamGoalEntities.Skip(batchCountIndex * TeamGoalsPerBatch).Take(TeamGoalsPerBatch);
                    foreach (var teamGoalEntity in teamGoalEntitiesBatch)
                    {
                        tableBatchOperation.Delete(teamGoalEntity);
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
                this.logger.LogError(ex, $"An error occurred in {nameof(this.DeleteTeamGoalDetailsAsync)} while deleting team goal details from storage.");
                throw;
            }
        }

        /// <summary>
        /// Get team goal details from storage depending upon filter condition.
        /// </summary>
        /// <param name="query">Query condition to fetch data from storage.</param>
        /// <returns>Returns collection of team goal details from storage.</returns>
        private async Task<IEnumerable<TeamGoalDetail>> GetTeamGoalDetailsAsync(TableQuery<TeamGoalDetail> query)
        {
            TableContinuationToken continuationToken = null;
            var teamGoalDetails = new List<TeamGoalDetail>();

            do
            {
                var queryResult = await this.CloudTable.ExecuteQuerySegmentedAsync(query, continuationToken);
                if (queryResult?.Results != null)
                {
                    teamGoalDetails.AddRange(queryResult.Results);
                    continuationToken = queryResult.ContinuationToken;
                }
            }
            while (continuationToken != null);

            return teamGoalDetails;
        }
    }
}
