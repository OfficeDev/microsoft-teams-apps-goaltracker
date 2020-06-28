// <copyright file="PersonalGoalSearchService.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Common
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Azure.Search;
    using Microsoft.Azure.Search.Models;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.GoalTracker.Models;

    /// <summary>
    /// Personal goal search service which will help in creating index, indexer and data source if it doesn't exist
    /// for indexing table which will be used for search align team goal.
    /// </summary>
    public class PersonalGoalSearchService : IPersonalGoalSearchService, IDisposable
    {
        /// <summary>
        /// Azure Search service maximum search result count for personal goal entity.
        /// </summary>
        private const int DefaultSearchResultCount = 1500;

        /// <summary>
        /// Azure Search service index name for personal goal details.
        /// </summary>
        private const string PersonalGoalIndexName = "personal-goal-index";

        /// <summary>
        /// Azure Search service indexer name for personal goal details.
        /// </summary>
        private const string PersonalGoalIndexerName = "personal-goal-indexer";

        /// <summary>
        /// Azure Search service data source name for personal goal details.
        /// </summary>
        private const string PersonalGoalDataSourceName = "personal-goal-storage";

        /// <summary>
        /// Table name where personal goal details are saved in storage.
        /// </summary>
        private const string PersonalGoalDetailTableName = "PersonalGoalDetail";

        /// <summary>
        /// Used to initialize task.
        /// </summary>
        private readonly Lazy<Task> initializeTask;

        /// <summary>
        /// Instance of Azure Search service client.
        /// </summary>
        private readonly SearchServiceClient searchServiceClient;

        /// <summary>
        /// Instance of Azure Search index client.
        /// </summary>
        private readonly SearchIndexClient searchIndexClient;

        /// <summary>
        /// Instance to send logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<PersonalGoalSearchService> logger;

        /// <summary>
        /// Represents a set of key/value Azure Search Service configuration properties.
        /// </summary>
        private readonly SearchServiceOptions searchServiceOptions;

        /// <summary>
        /// Flag: Has Dispose already been called?
        /// </summary>
        private bool disposed = false;

        /// <summary>
        /// Initializes a new instance of the <see cref="PersonalGoalSearchService"/> class.
        /// </summary>
        /// <param name="searchServiceOptions">A set of key/value Azure Search Service configuration properties.</param>
        /// <param name="storageOptions">A set of key/value application storage configuration properties.</param>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="searchServiceClient">Instance of Azure Search service client.</param>
        /// <param name="searchIndexClient">Instance of Azure Search index client.</param>
        public PersonalGoalSearchService(
            IOptions<SearchServiceOptions> searchServiceOptions,
            IOptionsMonitor<StorageOptions> storageOptions,
            ILogger<PersonalGoalSearchService> logger,
            SearchServiceClient searchServiceClient,
            SearchIndexClient searchIndexClient)
        {
            searchServiceOptions = searchServiceOptions ?? throw new ArgumentNullException(nameof(searchServiceOptions));
            storageOptions = storageOptions ?? throw new ArgumentNullException(nameof(storageOptions));

            this.searchServiceOptions = searchServiceOptions.Value;
            var searchServiceValue = this.searchServiceOptions.SearchServiceName;
            this.initializeTask = new Lazy<Task>(() => this.InitializeAsync(storageOptions.CurrentValue.ConnectionString));
            this.searchServiceClient = searchServiceClient;
            this.searchIndexClient = searchIndexClient;
            this.logger = logger;
        }

        /// <summary>
        /// Provide personal goal search result for personal goal detail table based on Azure Search service.
        /// </summary>
        /// <param name="searchScope">Scope of the search.</param>
        /// <param name="searchQuery">Query which the user had typed in Messaging Extension search field.</param>
        /// <param name="teamId">Unique identifier of team whose aligned goal status is requested.</param>
        /// <param name="count">Number of search results to return.</param>
        /// <param name="skip">Number of search results to skip.</param>
        /// <param name="sortBy">Represents sorting type like: Popularity or Newest.</param>
        /// <param name="filterQuery">Filter bar based query.</param>
        /// <returns>List of search results.</returns>
        public async Task<IEnumerable<TeamGoalStatus>> SearchPersonalGoalsWithStatusAsync(
            PersonalGoalSearchScope searchScope,
            string searchQuery,
            string teamId,
            int? count = null,
            int? skip = null,
            string sortBy = null,
            string filterQuery = null)
        {
            await this.EnsureInitializedAsync();
            IEnumerable<TeamGoalStatus> teamGoalStatusDetails = new List<TeamGoalStatus>();

            SearchParameters searchParameters = new SearchParameters()
            {
                Top = count ?? DefaultSearchResultCount,
                Skip = skip ?? 0,
                IncludeTotalResultCount = true,
                Facets = new List<string>() { "TeamGoalId,count:1000" },
                Select = new[] { "TeamGoalId" },
                SearchMode = SearchMode.All,
            };

            switch (searchScope)
            {
                case PersonalGoalSearchScope.NotStarted:
                    searchParameters.Filter = $"Status eq {(int)PersonalGoalStatus.NotStarted} and TeamId eq '{teamId}'";
                    break;

                case PersonalGoalSearchScope.InProgress:
                    searchParameters.Filter = $"Status eq {(int)PersonalGoalStatus.InProgress} and TeamId eq '{teamId}'";
                    break;

                case PersonalGoalSearchScope.Completed:
                    searchParameters.Filter = $"Status eq {(int)PersonalGoalStatus.Completed} and TeamId eq '{teamId}'";
                    break;
            }

            searchParameters.Filter += " and IsActive eq true and IsDeleted eq false";

            var goalResult = await this.searchIndexClient.Documents.SearchAsync<PersonalGoalDetail>(searchQuery, searchParameters);

            if (goalResult != null)
            {
                teamGoalStatusDetails = goalResult.Facets.Values.First()
                    .Select(facet => new TeamGoalStatus
                    {
                        TeamGoalId = facet.Value.ToString(),
                        NotStartedGoalCount = (searchScope == PersonalGoalSearchScope.NotStarted) ? Convert.ToInt32(facet.Count, CultureInfo.InvariantCulture) : 0,
                        InProgressGoalCount = (searchScope == PersonalGoalSearchScope.InProgress) ? Convert.ToInt32(facet.Count, CultureInfo.InvariantCulture) : 0,
                        CompletedGoalCount = (searchScope == PersonalGoalSearchScope.Completed) ? Convert.ToInt32(facet.Count, CultureInfo.InvariantCulture) : 0,
                    });
            }

            this.logger.LogInformation("Retrieved documents from goal search service successfully.");
            return teamGoalStatusDetails;
        }

        /// <summary>
        /// Creates Index, Data Source and Indexer for search service.
        /// </summary>
        /// <param name="storageConnectionString">Connection string to the data store.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        public async Task InitializeSearchServiceIndexAsync(string storageConnectionString)
        {
            try
            {
                await this.CreateSearchIndexAsync();
                await this.CreateDataSourceAsync(storageConnectionString);
                await this.CreateIndexerAsync();
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Dispose search service instance.
        /// </summary>
        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Initialization of InitializeAsync method which will help in indexing.
        /// </summary>
        /// <returns>Represents an asynchronous operation.</returns>
        public Task EnsureInitializedAsync()
        {
            return this.initializeTask.Value;
        }

        /// <summary>
        /// Protected implementation of Dispose pattern.
        /// </summary>
        /// <param name="disposing">True if already disposed else false.</param>
        protected virtual void Dispose(bool disposing)
        {
            if (this.disposed)
            {
                return;
            }

            if (disposing)
            {
                this.searchServiceClient.Dispose();
                this.searchIndexClient.Dispose();
            }

            this.disposed = true;
        }

        /// <summary>
        /// Create index, indexer and data source if doesn't exist.
        /// </summary>
        /// <param name="storageConnectionString">Connection string to the data store.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task InitializeAsync(string storageConnectionString)
        {
            try
            {
                await this.InitializeSearchServiceIndexAsync(storageConnectionString);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Failed to initialize Azure Search Service: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Create index in Azure Search service if it doesn't exist.
        /// </summary>
        /// <returns><see cref="Task"/> That represents index is created if it is not created.</returns>
        private async Task CreateSearchIndexAsync()
        {
            if (await this.searchServiceClient.Indexes.ExistsAsync(PersonalGoalIndexName))
            {
                await this.searchServiceClient.Indexes.DeleteAsync(PersonalGoalIndexName);
            }

            var tableIndex = new Index()
            {
                Name = PersonalGoalIndexName,
                Fields = FieldBuilder.BuildForType<PersonalGoalDetail>(),
            };
            await this.searchServiceClient.Indexes.CreateAsync(tableIndex);
        }

        /// <summary>
        /// Create data source if it doesn't exist in Azure Search service.
        /// </summary>
        /// <param name="connectionString">Connection string to the data store.</param>
        /// <returns><see cref="Task"/> That represents data source is added to Azure Search service.</returns>
        private async Task CreateDataSourceAsync(string connectionString)
        {
            if (await this.searchServiceClient.DataSources.ExistsAsync(PersonalGoalDataSourceName))
            {
                return;
            }

            var dataSource = DataSource.AzureTableStorage(
                name: PersonalGoalDataSourceName,
                storageConnectionString: connectionString,
                tableName: PersonalGoalDetailTableName,
                query: null,
                new SoftDeleteColumnDeletionDetectionPolicy("IsActive", false));

            await this.searchServiceClient.DataSources.CreateAsync(dataSource);
        }

        /// <summary>
        /// Create indexer if it doesn't exist in Azure Search service.
        /// </summary>
        /// <returns><see cref="Task"/> That represents indexer is created if not available in Azure Search service.</returns>
        private async Task CreateIndexerAsync()
        {
            try
            {
                if (await this.searchServiceClient.Indexers.ExistsAsync(PersonalGoalIndexerName))
                {
                    await this.searchServiceClient.Indexers.DeleteAsync(PersonalGoalIndexerName);
                }

                var indexer = new Indexer()
                {
                    Name = PersonalGoalIndexerName,
                    DataSourceName = PersonalGoalDataSourceName,
                    TargetIndexName = PersonalGoalIndexName,
                };

                await this.searchServiceClient.Indexers.CreateAsync(indexer);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Failed to create index: {ex.Message}");
                throw;
            }
        }
    }
}
