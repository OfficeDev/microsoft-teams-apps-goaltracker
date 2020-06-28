// <copyright file="GoalDeletionBackgroundService.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Common
{
    using System;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using Cronos;
    using Microsoft.Extensions.Hosting;
    using Microsoft.Extensions.Logging;

    /// <summary>
    /// BackgroundService class that inherits IHostedService and implements the methods related to background tasks for deleting personal and team goals once a week.
    /// </summary>
    public class GoalDeletionBackgroundService : BackgroundService
    {
        /// <summary>
        /// Instance to send logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<GoalDeletionBackgroundService> logger;

        /// <summary>
        /// Storage provider for working with personal goal data in storage.
        /// </summary>
        private readonly IPersonalGoalStorageProvider personalGoalStorageProvider;

        /// <summary>
        /// Storage provider for working with team goal data in storage.
        /// </summary>
        private readonly ITeamGoalStorageProvider teamGoalStorageProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="GoalDeletionBackgroundService"/> class.
        /// BackgroundService class that inherits IHostedService and implements the methods related to delete personal and team goals.
        /// </summary>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="personalGoalStorageProvider">Storage provider for working with personal goal data in storage.</param>
        /// <param name="teamGoalStorageProvider">Storage provider for working with team goal data in storage.</param>
        public GoalDeletionBackgroundService(
            ILogger<GoalDeletionBackgroundService> logger,
            IPersonalGoalStorageProvider personalGoalStorageProvider,
            ITeamGoalStorageProvider teamGoalStorageProvider)
        {
            this.logger = logger;
            this.personalGoalStorageProvider = personalGoalStorageProvider;
            this.teamGoalStorageProvider = teamGoalStorageProvider;
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
                    this.logger.LogInformation("Goal deletion background job execution has started.");
                    await this.DeletePersonalAndTeamGoalDetailsAsync();

                    // Schedule each run on Sunday
                    CronExpression goalDeletionCronExpression = CronExpression.Parse("0 0 * * SUN");
                    var next = goalDeletionCronExpression.GetNextOccurrence(DateTimeOffset.Now, TimeZoneInfo.Local);
                    var delay = next.HasValue ? next.Value - DateTimeOffset.Now : TimeSpan.FromDays(7);
                    await Task.Delay(delay, stoppingToken);
                }
#pragma warning disable CA1031 // Catching all generic exceptions in order to log exception details in logger and continue with next execution
                catch (Exception ex)
#pragma warning restore CA1031 // Catching all generic exceptions in order to log exception details in logger and continue with next execution
                {
                    this.logger.LogError(ex, $"Error while deleting goals at {nameof(this.DeletePersonalAndTeamGoalDetailsAsync)}: {ex}");
                    throw;
                }

                this.logger.LogInformation("Goal deletion background job execution has either stopped or did not execute.");
            }
        }

        /// <summary>
        /// Delete personal and team goal details where IsDeleted flag is true.
        /// </summary>
        /// <returns>A task that represents personal and team goal details need to be deleted from storage.</returns>
        private async Task DeletePersonalAndTeamGoalDetailsAsync()
        {
            var personalGoalDetails = await this.personalGoalStorageProvider.GetPersonalDeletedGoalDetailsAsync();
            if (personalGoalDetails != null && personalGoalDetails.Any())
            {
                await this.personalGoalStorageProvider.DeletePersonalGoalDetailsAsync(personalGoalDetails);
            }

            var teamGoalDetails = await this.teamGoalStorageProvider.GetDeletedTeamGoalDetailsAsync();
            if (teamGoalDetails != null && teamGoalDetails.Any())
            {
                await this.teamGoalStorageProvider.DeleteTeamGoalDetailsAsync(teamGoalDetails);
            }
        }
    }
}