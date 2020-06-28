// <copyright file="GoalBackgroundService.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Common
{
    using System;
    using System.Globalization;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Hosting;
    using Microsoft.Extensions.Logging;

    /// <summary>
    /// BackgroundService class that inherits IHostedService and implements the methods related to background tasks.
    /// Background service handles all the background tasks related to goals like send/update cards in team member personal bot.
    /// </summary>
    public class GoalBackgroundService : BackgroundService
    {
        private readonly BackgroundTaskWrapper taskWrapper;
        private readonly ILogger logger;
        private CancellationTokenSource tokenSource;
        private Task currentTask;

        /// <summary>
        /// Initializes a new instance of the <see cref="GoalBackgroundService"/> class.
        /// </summary>
        /// <param name="taskWrapper">Wrapper class instance for BackgroundTask.</param>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        public GoalBackgroundService(BackgroundTaskWrapper taskWrapper, ILogger<GoalBackgroundService> logger)
        {
            this.taskWrapper = taskWrapper;
            this.logger = logger;
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
            // Creating a linked token so that we can trigger cancellation outside of this token's cancellation.
            this.tokenSource = CancellationTokenSource.CreateLinkedTokenSource(stoppingToken);

            while (!stoppingToken.IsCancellationRequested)
            {
                try
                {
                    this.logger.LogInformation($"BackgroundService Dequeue method start at: {DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture)}");

                    // This is invoked when any task is enqueued in task wrapper.
                    // Dequeuing a task and running it in background until the cancellation is triggered or task is completed
                    this.currentTask = this.taskWrapper.DequeueAsync(this.tokenSource.Token);
                    await this.currentTask;
                }
#pragma warning disable CA1031 // Catching all generic exceptions in order to log exception details in logger and continue with next execution
                catch (Exception ex)
#pragma warning restore CA1031 // Catching all generic exceptions in order to log exception details in logger and continue with next execution
                {
                    // Execution has been canceled.
                    this.logger.LogError(ex, "Error while sending card to personal bot from background service.");
                }
            }

            this.logger.LogInformation("Goal reminder execution has either stopped or did not execute.");
        }
    }
}
