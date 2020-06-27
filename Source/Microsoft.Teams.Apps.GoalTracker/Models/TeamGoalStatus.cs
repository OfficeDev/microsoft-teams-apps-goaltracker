// <copyright file="TeamGoalStatus.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Models
{
    /// <summary>
    /// Class represents team goal status counts for not started, in progress and completed goals.
    /// </summary>
    public class TeamGoalStatus
    {
        /// <summary>
        /// Gets or sets unique identifier for team goal.
        /// </summary>
        public string TeamGoalId { get; set; }

        /// <summary>
        /// Gets or sets team goal name.
        /// </summary>
        public string TeamGoalName { get; set; }

        /// <summary>
        /// Gets or sets count of teams goals with not started status.
        /// </summary>
        public int? NotStartedGoalCount { get; set; }

        /// <summary>
        /// Gets or sets count of teams goals with in progress status.
        /// </summary>
        public int? InProgressGoalCount { get; set; }

        /// <summary>
        /// Gets or sets count of teams goals with completed status.
        /// </summary>
        public int? CompletedGoalCount { get; set; }
    }
}
