// <copyright file="PolicyNames.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Authentication
{
    /// <summary>
    /// This class lists the names of the custom authorization policies in the project.
    /// </summary>
    public static class PolicyNames
    {
        /// <summary>
        /// The name of the authorization policy, MustBePartOfTeamPolicy. Indicates that user must be a valid team member.
        /// </summary>
        public const string MustBePartOfTeamPolicy = "MustBePartOfTeamPolicy";
    }
}
