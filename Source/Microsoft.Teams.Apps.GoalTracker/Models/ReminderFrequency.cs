// <copyright file="ReminderFrequency.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Models
{
    /// <summary>
    /// Represents reminder frequency of goal.
    /// </summary>
    public enum ReminderFrequency
    {
        /// <summary>
        /// Represents reminder to be sent on each Monday.
        /// </summary>
        Weekly = 0,

        /// <summary>
        /// Represents reminder to be sent on 1st and 16th day of the month.
        /// </summary>
        Biweekly = 1,

        /// <summary>
        /// Represents reminder to be sent monthly.
        /// </summary>
        Monthly = 2,

        /// <summary>
        /// Represents reminder to be sent every quarter.
        /// </summary>
        Quarterly = 3,
    }
}
