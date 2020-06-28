// <copyright file="GoalReminderNotificationService.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Common
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using Cronos;
    using Microsoft.Extensions.Hosting;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.GoalTracker.Models;

    /// <summary>
    /// BackgroundService class that inherits IHostedService and implements the methods related to background tasks for sending notification once a day.
    /// </summary>
    public class GoalReminderNotificationService : BackgroundService
    {
        /// <summary>
        /// Goal cycle reminder notification to be sent three days prior to end date.
        /// </summary>
        private const int LookBackDays = 3;

        /// <summary>
        /// Instance to send logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<GoalReminderNotificationService> logger;

        /// <summary>
        /// Storage provider for working with personal goal data in storage.
        /// </summary>
        private readonly IPersonalGoalStorageProvider personalGoalStorageProvider;

        /// <summary>
        /// Storage provider for working with team goal data in storage.
        /// </summary>
        private readonly ITeamGoalStorageProvider teamGoalStorageProvider;

        /// <summary>
        /// Goal reminder activity helper to send reminder in team and personal scope.
        /// </summary>
        private readonly IGoalReminderActivityHelper goalReminderActivityHelper;

        /// <summary>
        /// Initializes a new instance of the <see cref="GoalReminderNotificationService"/> class.
        /// BackgroundService class that inherits IHostedService and implements the methods related to sending notification tasks.
        /// </summary>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="personalGoalStorageProvider">Storage provider for working with personal goal data in storage.</param>
        /// <param name="teamGoalStorageProvider">Storage provider for working with team goal data in storage.</param>
        /// <param name="goalReminderActivityHelper">Goal reminder activity helper to send reminder in team and personal scope.</param>
        public GoalReminderNotificationService(
            ILogger<GoalReminderNotificationService> logger,
            IPersonalGoalStorageProvider personalGoalStorageProvider,
            ITeamGoalStorageProvider teamGoalStorageProvider,
            IGoalReminderActivityHelper goalReminderActivityHelper)
        {
            this.logger = logger;
            this.personalGoalStorageProvider = personalGoalStorageProvider;
            this.teamGoalStorageProvider = teamGoalStorageProvider;
            this.goalReminderActivityHelper = goalReminderActivityHelper;
        }

        /// <summary>
        /// This method is called when the Microsoft.Extensions.Hosting.IHostedService starts.
        /// The implementation should return a task that represents the lifetime of the long
        /// running operation(s) being performed.
        /// </summary>
        /// <param name="stoppingToken">Triggered when Microsoft.Extensions.Hosting.IHostedService. StopAsync(System.Threading.CancellationToken) is called.</param>
        /// <returns>A System.Threading.Tasks.Task that represents the long running operations.</returns>
        protected async override Task ExecuteAsync(CancellationToken stoppingToken)
        {
            while (!stoppingToken.IsCancellationRequested)
            {
                try
                {
                    this.logger.LogInformation("Goal reminder background job execution has started.");
                    await this.ProcessGoalReminderAsync();

                    // Schedule each run at midnight
                    CronExpression goalReminderCronExpression = CronExpression.Parse("0 0 */1 * *");
                    var next = goalReminderCronExpression.GetNextOccurrence(DateTimeOffset.Now, TimeZoneInfo.Local);
                    var delay = next.HasValue ? next.Value - DateTimeOffset.Now : TimeSpan.FromDays(1);
                    await Task.Delay(delay, stoppingToken);
                }
#pragma warning disable CA1031 // Catching all generic exceptions in order to log exception details in logger and continue with next execution
                catch (Exception ex)
#pragma warning restore CA1031 // Catching all generic exceptions in order to log exception details in logger and continue with next execution
                {
                    this.logger.LogError(ex, $"Error while sending goal reminder card at {nameof(this.ProcessGoalReminderAsync)}: {ex}");
                }
            }

            this.logger.LogInformation("Goal reminder execution has either stopped or did not execute.");
        }

        /// <summary>
        /// Process today's personal and team goal details from storage and send the reminder according to frequency or end date.
        /// </summary>
        /// <returns>A task that schedules goal reminder notification for personal and team goals.</returns>
        private async Task ProcessGoalReminderAsync()
        {
            var personalGoalReminderDetails = await this.GetPersonalGoalReminderDetailsAsync();
            if (personalGoalReminderDetails != null && personalGoalReminderDetails.Any())
            {
                // Process goal reminder for the bot, installed in the personal scope.
                await this.SendGoalReminderToPersonalBotAsync(personalGoalReminderDetails);
            }

            var teamGoalReminderDetails = await this.GetTeamGoalReminderDetailsAsync();
            if (teamGoalReminderDetails != null && teamGoalReminderDetails.Any())
            {
                // Process goal reminder for team and team members.
                await this.SendGoalReminderToTeamAndTeamMembersAsync(teamGoalReminderDetails);
            }
        }

        /// <summary>
        /// Get personal goal reminder details from storage.
        /// </summary>
        /// <returns>Collection of personal goal reminder details.</returns>
        private async Task<IEnumerable<PersonalGoalDetail>> GetPersonalGoalReminderDetailsAsync()
        {
            var today = DateTime.UtcNow.ToString(Constants.UTCDateFormat, CultureInfo.InvariantCulture);
            var allPersonalUnalignedGoals = await this.personalGoalStorageProvider.GetPersonalUnalignedGoalReminderDetailsAsync();
            if (!allPersonalUnalignedGoals.Any())
            {
                return null;
            }

            List<PersonalGoalDetail> personalGoalReminderDetails = new List<PersonalGoalDetail>();
            foreach (var personalGoalDetail in allPersonalUnalignedGoals)
            {
                // Check whether personal goal detail already exists in personalGoalReminderDetails i.e. unique row will be picked for each user.
                var personalGoal = personalGoalReminderDetails.Where(currentPersonalGoal => currentPersonalGoal.UserAadObjectId == personalGoalDetail.UserAadObjectId).FirstOrDefault();
                if (personalGoal != null)
                {
                    continue;
                }

                var personalGoalEndDate = DateTime.Parse(personalGoalDetail.EndDateUTC, CultureInfo.InvariantCulture).AddDays(+1).ToString(Constants.UTCDateFormat, CultureInfo.InvariantCulture);
                if (personalGoalEndDate == today)
                {
                    // Unaligned goals cycle is completed, no need to send the reminder and update personal goal detail table
                    await this.goalReminderActivityHelper.UpdatePersonalGoalAndNoteDetailsAsync(personalGoalDetail.UserAadObjectId);
                }
                else
                {
                    // Sending reminders only for unaligned goals, for aligned goals reminder will be sent from team where user has aligned his/her goal.
                    personalGoalReminderDetails.Add(personalGoalDetail);
                }
            }

            return personalGoalReminderDetails;
        }

        /// <summary>
        /// Send goal reminder according to reminder frequency or end date in comparison with today's date for personal goals.
        /// </summary>
        /// <param name="personalGoalReminderDetails">Personal goal reminder details.</param>
        /// <remarks>The application is designed to send personal notifications while iterating at app level.
        /// This would work for smaller orgarnization where notifications up to 5K at a time might work.
        /// Tip: Try to queue the notification requests using Azure service bus and fetch in order to handle bot throttling limit.
        /// Read more about bot throttling limit : https://docs.microsoft.com/en-us/microsoftteams/platform/bots/how-to/rate-limit </remarks>
        /// <returns>A task that sends goal reminder card to the bot, installed in the personal scope.</returns>
        private async Task SendGoalReminderToPersonalBotAsync(IEnumerable<PersonalGoalDetail> personalGoalReminderDetails)
        {
            var today = DateTime.UtcNow.ToString(Constants.UTCDateFormat, CultureInfo.InvariantCulture);
            foreach (var personalGoalDetail in personalGoalReminderDetails)
            {
                var personalGoalEndDate = DateTime.Parse(personalGoalDetail.EndDateUTC, CultureInfo.InvariantCulture).AddDays(-LookBackDays).ToString(Constants.UTCDateFormat, CultureInfo.InvariantCulture);
                if (personalGoalEndDate == today)
                {
                    // Send goal reminder if today's date is before three days to end date.
                    await this.goalReminderActivityHelper.SendGoalReminderToPersonalBotAsync(personalGoalDetail, isReminderBeforeThreeDays: true);
                }
                else
                {
                    // Send goal reminder according to reminder frequency.
                    await this.goalReminderActivityHelper.SendGoalReminderToPersonalBotAsync(personalGoalDetail);
                }
            }
        }

        /// <summary>
        /// Get team goal reminder details from storage.
        /// </summary>
        /// <returns>Collection of team goal reminder details.</returns>
        private async Task<IEnumerable<TeamGoalDetail>> GetTeamGoalReminderDetailsAsync()
        {
            var today = DateTime.UtcNow.ToString(Constants.UTCDateFormat, CultureInfo.InvariantCulture);
            var allTeamGoalReminderDetails = await this.teamGoalStorageProvider.GetTeamGoalReminderDetailsAsync();
            if (!allTeamGoalReminderDetails.Any())
            {
                return null;
            }

            List<TeamGoalDetail> teamGoalReminderDetails = new List<TeamGoalDetail>();
            foreach (var teamGoalDetail in allTeamGoalReminderDetails)
            {
                // Check whether team goal detail already exists in teamGoalReminderDetails i.e. unique row will be picked for each team.
                var teamGoal = teamGoalReminderDetails.Where(currentTeamGoal => currentTeamGoal.TeamId == teamGoalDetail.TeamId).FirstOrDefault();
                if (teamGoal != null)
                {
                    continue;
                }

                var teamGoalEndDate = DateTime.Parse(teamGoalDetail.TeamGoalEndDateUTC, CultureInfo.InvariantCulture).AddDays(+1).ToString(Constants.UTCDateFormat, CultureInfo.InvariantCulture);
                if (teamGoalEndDate == today)
                {
                    // Team goal cycle is completed. No need to send the reminder, updating personal goal, team goal, personal goal note details in storage.
                    await this.goalReminderActivityHelper.UpdateGoalDetailsAsync(teamGoalDetail);
                }
                else
                {
                    teamGoalReminderDetails.Add(teamGoalDetail);
                }
            }

            return teamGoalReminderDetails;
        }

        /// <summary>
        /// Send goal reminder according to reminder frequency or end date in comparison with today's date for team goals.
        /// </summary>
        /// <param name="teamGoalReminderDetails">Team goal reminder details.</param>
        /// <returns>A task that sends goal reminder card to team and team members.</returns>
        private async Task SendGoalReminderToTeamAndTeamMembersAsync(IEnumerable<TeamGoalDetail> teamGoalReminderDetails)
        {
            var today = DateTime.UtcNow.ToString(Constants.UTCDateFormat, CultureInfo.InvariantCulture);
            foreach (var teamGoalDetail in teamGoalReminderDetails)
            {
                var teamGoalEndDate = DateTime.Parse(teamGoalDetail.TeamGoalEndDateUTC, CultureInfo.InvariantCulture).AddDays(-LookBackDays).ToString(Constants.UTCDateFormat, CultureInfo.InvariantCulture);
                if (teamGoalEndDate == today)
                {
                    // Send goal reminder if today's date is before three days to end date.
                    await this.goalReminderActivityHelper.SendGoalReminderToTeamAndTeamMembersAsync(teamGoalDetail, isReminderBeforeThreeDays: true);
                }
                else
                {
                    // Send goal reminder according to reminder frequency.
                    await this.goalReminderActivityHelper.SendGoalReminderToTeamAndTeamMembersAsync(teamGoalDetail);
                }
            }
        }
    }
}
