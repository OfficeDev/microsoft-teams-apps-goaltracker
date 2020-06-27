// <copyright file="BaseStorageProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Common
{
    using System;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.GoalTracker.Models;
    using Microsoft.WindowsAzure.Storage;
    using Microsoft.WindowsAzure.Storage.RetryPolicies;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Implements storage provider which initializes table if not exists and provide table client instance.
    /// </summary>
    public class BaseStorageProvider
    {
        /// <summary>
        /// Storage connection string where all tables will be created and data will be stored.
        /// </summary>
        private readonly string connectionString;

        /// <summary>
        /// Initializes a new instance of the <see cref="BaseStorageProvider"/> class.
        /// </summary>
        /// <param name="storageOptions">A set of key/value application storage configuration properties.</param>
        /// <param name="tableName">Table name of storage to initialize.</param>
        public BaseStorageProvider(IOptionsMonitor<StorageOptions> storageOptions, string tableName)
        {
            storageOptions = storageOptions ?? throw new ArgumentNullException(nameof(storageOptions));
            this.connectionString = storageOptions.CurrentValue.ConnectionString;
            this.TableName = tableName;
            this.InitializeTask = new Lazy<Task>(() => this.InitializeAsync());
        }

        /// <summary>
        /// Gets or sets a lazy task to initialize storage.
        /// </summary>
        protected Lazy<Task> InitializeTask { get; set; }

        /// <summary>
        /// Gets or sets Microsoft Azure Table storage table name.
        /// </summary>
        protected string TableName { get; set; }

        /// <summary>
        /// Gets or sets a table in the Microsoft Azure Table storage.
        /// </summary>
        protected CloudTable CloudTable { get; set; }

        /// <summary>
        /// Ensures Microsoft Azure Table storage should be created before working on table.
        /// </summary>
        /// <returns>Represents an asynchronous operation.</returns>
        protected async Task EnsureInitializedAsync()
        {
            await this.InitializeTask.Value;
        }

        /// <summary>
        /// Create storage table if it does not exist.
        /// </summary>
        /// <returns><see cref="Task"/> representing the asynchronous operation task which represents table is created if it does not exists.</returns>
        private async Task<CloudTable> InitializeAsync()
        {
            // Exponential retry policy with back off of 0.5 seconds and 3 retries.
            var exponentialRetryPolicy = new TableRequestOptions()
            {
                RetryPolicy = new ExponentialRetry(TimeSpan.FromSeconds(0.5), 3),
            };

            CloudStorageAccount storageAccount = CloudStorageAccount.Parse(this.connectionString);
            CloudTableClient cloudTableClient = storageAccount.CreateCloudTableClient();
            cloudTableClient.DefaultRequestOptions = exponentialRetryPolicy;
            this.CloudTable = cloudTableClient.GetTableReference(this.TableName);
            await this.CloudTable.CreateIfNotExistsAsync();

            return this.CloudTable;
        }
    }
}