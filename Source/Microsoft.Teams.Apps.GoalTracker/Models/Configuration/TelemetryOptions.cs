// <copyright file="TelemetryOptions.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Models
{
    /// <summary>
    /// Provides application setting related to application insights.
    /// </summary>
    public class TelemetryOptions
    {
        /// <summary>
        /// Gets or sets the Instrumentation key of application insights.
        /// </summary>
        public string InstrumentationKey { get; set; }
    }
}
