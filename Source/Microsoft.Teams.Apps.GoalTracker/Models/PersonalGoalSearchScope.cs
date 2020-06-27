// <copyright file="PersonalGoalSearchScope.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Models
{
    /// <summary>
    /// Class represents scope on basis of which goals will be searched in goal status command.
    /// </summary>
    public enum PersonalGoalSearchScope
    {
        /// <summary>
        /// Goals which are not yet started.
        /// </summary>
        NotStarted,

        /// <summary>
        /// Goals which are started and are in progress.
        /// </summary>
        InProgress,

        /// <summary>
        /// Goals which are completed.
        /// </summary>
        Completed,
    }
}
