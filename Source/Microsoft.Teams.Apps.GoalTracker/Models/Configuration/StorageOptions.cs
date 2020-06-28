// <copyright file="StorageOptions.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Models
{
    /// <summary>
    /// Provides application setting related to storage.
    /// </summary>
    public class StorageOptions
    {
        /// <summary>
        /// Gets or sets storage connection string where all tables will be created and data will be stored..
        /// </summary>
        public string ConnectionString { get; set; }
    }
}
