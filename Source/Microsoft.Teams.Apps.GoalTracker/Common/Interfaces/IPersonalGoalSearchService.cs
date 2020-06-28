// <copyright file="IPersonalGoalSearchService.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Common
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.GoalTracker.Models;

    /// <summary>
    /// Interface for personal goal search service which helps in searching personal goals using Azure Search service.
    /// </summary>
    public interface IPersonalGoalSearchService
    {
        /// <summary>
        /// Provide search result for table to be used by user's based on Azure Search service.
        /// </summary>
        /// <param name="searchScope">Scope of the search.</param>
        /// <param name="searchQuery">Query which the user had typed in Messaging Extension search field.</param>
        /// <param name="teamId">Unique identifier of team whose aligned goal status is requested.</param>
        /// <param name="count">Number of search results to return.</param>
        /// <param name="skip">Number of search results to skip.</param>
        /// <param name="sortBy">Represents sorting type like: Popularity or Newest.</param>
        /// <param name="filterQuery">Filter bar based query.</param>
        /// <returns>List of search results.</returns>
        Task<IEnumerable<TeamGoalStatus>> SearchPersonalGoalsWithStatusAsync(
            PersonalGoalSearchScope searchScope,
            string searchQuery,
            string teamId,
            int? count = null,
            int? skip = null,
            string sortBy = null,
            string filterQuery = null);

        /// <summary>
        /// Creates Index, Data Source and Indexer for search service.
        /// </summary>
        /// <param name="storageConnectionString">Connection string to the data store.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        Task InitializeSearchServiceIndexAsync(string storageConnectionString);

        /// <summary>
        /// Initialization of InitializeAsync method which will help in indexing.
        /// </summary>
        /// <returns>Represents an asynchronous operation.</returns>
        Task EnsureInitializedAsync();
    }
}
