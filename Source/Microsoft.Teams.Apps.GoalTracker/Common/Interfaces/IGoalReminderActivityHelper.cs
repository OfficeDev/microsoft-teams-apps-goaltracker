// <copyright file="IGoalReminderActivityHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Common
{
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.GoalTracker.Models;

    /// <summary>
    /// Handles goal reminders to be sent to team or personal bot depending on the goal cycle end date.
    /// </summary>
    public interface IGoalReminderActivityHelper
    {
        /// <summary>
        /// Method to send goal reminder card to personal bot.
        /// </summary>
        /// <param name="personalGoalDetail">Holds personal goal detail entity data sent from background service.</param>
        /// <param name="isReminderBeforeThreeDays">Determines reminder to be sent prior 3 days to end date.</param>
        /// <returns>A Task represents goal reminder card is sent to the bot, installed in the personal scope.</returns>
        Task SendGoalReminderToPersonalBotAsync(PersonalGoalDetail personalGoalDetail, bool isReminderBeforeThreeDays = false);

        /// <summary>
        /// Update personal goal details and personal goal note details in storage when personal goal cycle is ended.
        /// </summary>
        /// <param name="userAadObjectId">AAD object id of the user whose goal details needs to be deleted.</param>
        /// <returns>A task that represents personal goal detail entity data is saved or updated.</returns>
        Task UpdatePersonalGoalAndNoteDetailsAsync(string userAadObjectId);

        /// <summary>
        /// Method to send goal reminder card to team and team members.
        /// </summary>
        /// <param name="teamGoalDetail">Holds team goal detail entity data sent from background service.</param>
        /// <param name="isReminderBeforeThreeDays">Determines reminder to be sent prior 3 days to end date.</param>
        /// <returns>A Task represents goal reminder card is sent to team and team members.</returns>
        Task SendGoalReminderToTeamAndTeamMembersAsync(TeamGoalDetail teamGoalDetail, bool isReminderBeforeThreeDays = false);

        /// <summary>
        /// Method to update team goal, personal goal and note details in storage when team goal is ended.
        /// </summary>
        /// <param name="teamGoalDetail">Holds team goal detail entity data sent from background service.</param>
        /// <returns>A Task represents goal reminder card is sent to team and team members.</returns>
        Task UpdateGoalDetailsAsync(TeamGoalDetail teamGoalDetail);
    }
}
