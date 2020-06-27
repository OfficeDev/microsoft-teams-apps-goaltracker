// <copyright file="GoalTrackerActivityHandlerOptions.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Models
{
    /// <summary>
    /// The GoalTrackerActivityHandlerOptions are the options for the <see cref="GoalTrackerActivityHandlerOptions" /> bot.
    /// </summary>
    public sealed class GoalTrackerActivityHandlerOptions
    {
        /// <summary>
        /// Gets or sets unique id of Tenant.
        /// </summary>
        public string TenantId { get; set; }

        /// <summary>
        /// Gets or sets application base Uri.
        /// </summary>
        public string AppBaseUri { get; set; }

        /// <summary>
        /// Gets or sets entity id of static goals tab.
        /// </summary>
        public string GoalsTabEntityId { get; set; }

        /// <summary>
        /// Gets or sets application manifest id.
        /// </summary>
        public string ManifestId { get; set; }
    }
}