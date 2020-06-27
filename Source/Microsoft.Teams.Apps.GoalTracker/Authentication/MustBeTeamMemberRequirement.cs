// <copyright file="MustBeTeamMemberRequirement.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Authentication
{
    using Microsoft.AspNetCore.Authorization;

    /// <summary>
    /// This authorization class implements the marker interface
    /// <see cref="IAuthorizationRequirement"/> to check if user meets teams member specific requirements
    /// for accessing resources.
    /// It specifies that a user is a member of a certain team.
    /// </summary>
    public class MustBeTeamMemberRequirement : IAuthorizationRequirement
    {
    }
}
