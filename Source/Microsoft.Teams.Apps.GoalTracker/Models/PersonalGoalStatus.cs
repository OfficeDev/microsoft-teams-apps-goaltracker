// <copyright file="PersonalGoalStatus.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Models
{
    /// <summary>
    /// Represents the current status of a personal goal.
    /// </summary>
    public enum PersonalGoalStatus
    {
        /// <summary>
        /// Represents a goal which is not yet started.
        /// </summary>
        NotStarted = 0,

        /// <summary>
        /// Represents a goal which is in progress.
        /// </summary>
        InProgress = 1,

        /// <summary>
        /// Represents a goal which is completed.
        /// </summary>
        Completed = 2,
    }
}
