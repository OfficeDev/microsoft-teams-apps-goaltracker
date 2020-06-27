// <copyright file="TeamStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Common
{
    using System;
    using System.Net;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.GoalTracker.Models;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Implements storage provider which helps in storing, updating, deleting team details in which bot is installed.
    /// </summary>
    public class TeamStorageProvider : BaseStorageProvider, ITeamStorageProvider
    {
        private const string TeamDetailTable = "TeamDetail";

        /// <summary>
        /// Initializes a new instance of the <see cref="TeamStorageProvider"/> class.
        /// </summary>
        /// <param name="storageOptions">A set of key/value application storage configuration properties.</param>
        public TeamStorageProvider(IOptionsMonitor<StorageOptions> storageOptions)
            : base(storageOptions, TeamDetailTable)
        {
        }

        /// <summary>
        /// Store or update team detail in Azure table storage.
        /// </summary>
        /// <param name="teamEntity">Represents team entity used for storage and retrieval.</param>
        /// <returns><see cref="Task"/> that represents team entity is saved or updated.</returns>
        public async Task<bool> StoreOrUpdateTeamDetailAsync(TeamDetail teamEntity)
        {
            await this.EnsureInitializedAsync();
            teamEntity = teamEntity ?? throw new ArgumentNullException(nameof(teamEntity));
            TableOperation addOrUpdateOperation = TableOperation.InsertOrReplace(teamEntity);
            var result = await this.CloudTable.ExecuteAsync(addOrUpdateOperation);
            return result.HttpStatusCode == (int)HttpStatusCode.NoContent;
        }

        /// <summary>
        /// Get already team detail from Azure table storage.
        /// </summary>
        /// <param name="teamId">Team Id.</param>
        /// <returns><see cref="Task"/> Already saved team detail.</returns>
        public async Task<TeamDetail> GetTeamDetailAsync(string teamId)
        {
            await this.EnsureInitializedAsync();
            var operation = TableOperation.Retrieve<TeamDetail>(teamId, teamId);
            var data = await this.CloudTable.ExecuteAsync(operation);
            return data.Result as TeamDetail;
        }

        /// <summary>
        /// This method delete the team detail record from table.
        /// </summary>
        /// <param name="teamEntity">Team configuration table entity.</param>
        /// <returns>A <see cref="Task"/> of type bool where true represents entity record is successfully deleted from table while false indicates failure in deleting data.</returns>
        public async Task<bool> DeleteTeamDetailAsync(TeamDetail teamEntity)
        {
            await this.EnsureInitializedAsync();
            TableOperation insertOrMergeOperation = TableOperation.Delete(teamEntity);
            TableResult result = await this.CloudTable.ExecuteAsync(insertOrMergeOperation);
            return result.HttpStatusCode == (int)HttpStatusCode.NoContent;
        }
    }
}
