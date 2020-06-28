// <copyright file="GoalReminderActivityHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Helpers
{
    using System;
    using System.Globalization;
    using System.Linq;
    using System.Net;
    using System.Threading;
    using System.Threading.Tasks;
    using AdaptiveCards;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Integration.AspNet.Core;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Bot.Schema;
    using Microsoft.Extensions.Localization;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.GoalTracker.Cards;
    using Microsoft.Teams.Apps.GoalTracker.Common;
    using Microsoft.Teams.Apps.GoalTracker.Models;
    using Polly;
    using Polly.Contrib.WaitAndRetry;
    using Polly.Retry;

    /// <summary>
    /// Class to send goal reminder in personal bot and in team.
    /// </summary>
    public class GoalReminderActivityHelper : IGoalReminderActivityHelper
    {
        /// <summary>
        /// Represents retry delay.
        /// </summary>
        private const int RetryDelay = 1000;

        /// <summary>
        /// Represents retry count.
        /// </summary>
        private const int RetryCount = 2;

        /// <summary>
        /// Retry policy with jitter, retry twice with a jitter delay of up to 1 sec. Retry for HTTP 429(transient error)/502 bad gateway.
        /// </summary>
        /// <remarks>
        /// Reference: https://github.com/Polly-Contrib/Polly.Contrib.WaitAndRetry#new-jitter-recommendation.
        /// </remarks>
        private readonly AsyncRetryPolicy retryPolicy = Policy.Handle<ErrorResponseException>(
            ex => ex.Response.StatusCode == HttpStatusCode.TooManyRequests || ex.Response.StatusCode == HttpStatusCode.BadGateway)
            .WaitAndRetryAsync(Backoff.DecorrelatedJitterBackoffV2(TimeSpan.FromMilliseconds(RetryDelay), RetryCount));

        /// <summary>
        /// Microsoft application credentials.
        /// </summary>
        private readonly MicrosoftAppCredentials microsoftAppCredentials;

        /// <summary>
        /// Bot adapter.
        /// </summary>
        private readonly IBotFrameworkHttpAdapter adapter;

        /// <summary>
        /// Instance to send logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<GoalReminderActivityHelper> logger;

        /// <summary>
        /// The current cultures' string localizer.
        /// </summary>
        private readonly IStringLocalizer<Strings> localizer;

        /// <summary>
        /// A set of key/value application configuration properties for Activity settings.
        /// </summary>
        private readonly IOptions<GoalTrackerActivityHandlerOptions> options;

        /// <summary>
        /// Storage provider for working with personal goal data in storage.
        /// </summary>
        private readonly IPersonalGoalStorageProvider personalGoalStorageProvider;

        /// <summary>
        /// Storage provider for working with personal goal note data in storage.
        /// </summary>
        private readonly IPersonalGoalNoteStorageProvider personalGoalNoteStorageProvider;

        /// <summary>
        /// Storage provider for working with team goal data in storage.
        /// </summary>
        private readonly ITeamGoalStorageProvider teamGoalStorageProvider;

        /// <summary>
        /// Instance of class that handles goal helper methods.
        /// </summary>
        private readonly GoalHelper goalHelper;

        /// <summary>
        /// Initializes a new instance of the <see cref="GoalReminderActivityHelper"/> class.
        /// BackgroundService class that inherits IHostedService and implements the methods related to sending notification tasks.
        /// </summary>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="options">A set of key/value application configuration properties for activity handler.</param>
        /// <param name="microsoftAppCredentials">Instance for Microsoft application credentials.</param>
        /// <param name="adapter">An instance of bot adapter.</param>
        /// <param name="personalGoalStorageProvider">Storage provider for working with personal goal data in storage.</param>
        /// <param name="personalGoalNoteStorageProvider">Storage provider for working with personal goal note data in storage</param>
        /// <param name="teamGoalStorageProvider">Storage provider for working with team goal data in storage.</param>
        /// <param name="goalHelper">Instance of class that handles goal helper methods.</param>
        public GoalReminderActivityHelper(
            ILogger<GoalReminderActivityHelper> logger,
            IStringLocalizer<Strings> localizer,
            IOptions<GoalTrackerActivityHandlerOptions> options,
            MicrosoftAppCredentials microsoftAppCredentials,
            IBotFrameworkHttpAdapter adapter,
            IPersonalGoalStorageProvider personalGoalStorageProvider,
            IPersonalGoalNoteStorageProvider personalGoalNoteStorageProvider,
            ITeamGoalStorageProvider teamGoalStorageProvider,
            GoalHelper goalHelper)
        {
            this.logger = logger;
            this.localizer = localizer;
            this.options = options;
            this.microsoftAppCredentials = microsoftAppCredentials;
            this.adapter = adapter;
            this.personalGoalStorageProvider = personalGoalStorageProvider;
            this.personalGoalNoteStorageProvider = personalGoalNoteStorageProvider;
            this.teamGoalStorageProvider = teamGoalStorageProvider;
            this.goalHelper = goalHelper;
        }

        /// <summary>
        /// Method to send goal reminder card to personal bot.
        /// </summary>
        /// <param name="personalGoalDetail">Holds personal goal detail entity data sent from background service.</param>
        /// <param name="isReminderBeforeThreeDays">Determines reminder to be sent prior 3 days to end date.</param>
        /// <returns>A Task represents goal reminder card is sent to the bot, installed in the personal scope.</returns>
        public async Task SendGoalReminderToPersonalBotAsync(PersonalGoalDetail personalGoalDetail, bool isReminderBeforeThreeDays = false)
        {
            personalGoalDetail = personalGoalDetail ?? throw new ArgumentNullException(nameof(personalGoalDetail));
            var conversationReference = new ConversationReference()
            {
                ChannelId = Constants.TeamsBotFrameworkChannelId,
                Bot = new ChannelAccount() { Id = $"28:{this.microsoftAppCredentials.MicrosoftAppId}" },
                ServiceUrl = personalGoalDetail.ServiceUrl,
                Conversation = new ConversationAccount() { ConversationType = Constants.PersonalConversationType, Id = personalGoalDetail.ConversationId, TenantId = this.options.Value.TenantId },
            };

            try
            {
                await this.retryPolicy.ExecuteAsync(async () =>
                {
                    await ((BotFrameworkAdapter)this.adapter).ContinueConversationAsync(
                        this.microsoftAppCredentials.MicrosoftAppId,
                        conversationReference,
                        async (turnContext, cancellationToken) =>
                        {
                            string reminderType = string.Empty;
                            AdaptiveTextColor reminderTypeColor = AdaptiveTextColor.Accent;
                            if (isReminderBeforeThreeDays)
                            {
                                reminderType = this.localizer.GetString("PersonalGoalEndingAfterThreeDays");
                                reminderTypeColor = AdaptiveTextColor.Attention;
                            }
                            else
                            {
                                reminderType = this.GetReminderTypeString(personalGoalDetail.ReminderFrequency);
                            }

                            var goalReminderAttachment = MessageFactory.Attachment(GoalReminderCard.GetGoalReminderCard(this.localizer, this.options.Value.ManifestId, this.options.Value.GoalsTabEntityId, reminderType, reminderTypeColor));
                            this.logger.LogInformation($"Sending goal reminder card to Personal bot. Conversation id: {personalGoalDetail.ConversationId}");
                            await turnContext.SendActivityAsync(goalReminderAttachment, cancellationToken);
                        },
                        CancellationToken.None);
                });
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error while sending goal reminder card to personal bot: {ex.Message} at {nameof(this.SendGoalReminderToPersonalBotAsync)}");
                throw;
            }
        }

        /// <summary>
        /// Update personal goal details and personal goal note details in storage when personal goal cycle is ended.
        /// </summary>
        /// <param name="userAadObjectId">AAD object id of the user whose goal details needs to be deleted.</param>
        /// <returns>A task that represents personal goal detail entity data is saved or updated.</returns>
        public async Task UpdatePersonalGoalAndNoteDetailsAsync(string userAadObjectId)
        {
            try
            {
                var personalGoalEntities = await this.personalGoalStorageProvider.GetPersonalGoalDetailsByUserAadObjectIdAsync(userAadObjectId);
                personalGoalEntities = personalGoalEntities.Where(personalGoal => !personalGoal.IsAligned);

                // Update personal goal details data when goal cycle is ended for unaligned goals
                foreach (var personalGoalEntity in personalGoalEntities)
                {
                    personalGoalEntity.IsActive = false;
                    personalGoalEntity.IsReminderActive = false;
                    personalGoalEntity.LastModifiedOn = DateTime.UtcNow.ToString(CultureInfo.CurrentCulture);

                    var personalGoalNoteEntities = await this.personalGoalNoteStorageProvider.GetPersonalGoalNoteDetailsAsync(personalGoalEntity.PersonalGoalId, userAadObjectId);

                    // Update personal goal note details data when goal cycle is ended for unaligned goals
                    foreach (var personalGoalNoteEntity in personalGoalNoteEntities)
                    {
                        personalGoalNoteEntity.IsActive = false;
                        personalGoalNoteEntity.LastModifiedOn = DateTime.UtcNow.ToString(CultureInfo.CurrentCulture);
                    }

                    await this.personalGoalNoteStorageProvider.CreateOrUpdatePersonalGoalNoteDetailsAsync(personalGoalNoteEntities);
                }

                await this.personalGoalStorageProvider.CreateOrUpdatePersonalGoalDetailsAsync(personalGoalEntities);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Failed to save personal goal detail data to table storage at {nameof(this.UpdatePersonalGoalAndNoteDetailsAsync)}: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Method to send goal reminder card to team and team members or update team goal, personal goal and note details in storage.
        /// </summary>
        /// <param name="teamGoalDetail">Holds team goal detail entity data sent from background service.</param>
        /// <param name="isReminderBeforeThreeDays">Determines reminder to be sent prior 3 days to end date.</param>
        /// <returns>A Task represents goal reminder card is sent to team and team members.</returns>
        public async Task SendGoalReminderToTeamAndTeamMembersAsync(TeamGoalDetail teamGoalDetail, bool isReminderBeforeThreeDays = false)
        {
            teamGoalDetail = teamGoalDetail ?? throw new ArgumentNullException(nameof(teamGoalDetail));

            try
            {
                var teamId = teamGoalDetail.TeamId;
                string serviceUrl = teamGoalDetail.ServiceUrl;
                MicrosoftAppCredentials.TrustServiceUrl(serviceUrl);

                var conversationReference = new ConversationReference()
                {
                    ChannelId = Constants.TeamsBotFrameworkChannelId,
                    Bot = new ChannelAccount() { Id = $"28:{this.microsoftAppCredentials.MicrosoftAppId}" },
                    ServiceUrl = serviceUrl,
                    Conversation = new ConversationAccount() { ConversationType = Constants.ChannelConversationType, IsGroup = true, Id = teamId, TenantId = this.options.Value.TenantId },
                };

                await this.retryPolicy.ExecuteAsync(async () =>
                {
                    try
                    {
                        await ((BotFrameworkAdapter)this.adapter).ContinueConversationAsync(
                            this.microsoftAppCredentials.MicrosoftAppId,
                            conversationReference,
                            async (turnContext, cancellationToken) =>
                            {
                                string reminderType = string.Empty;
                                AdaptiveTextColor reminderTypeColor = AdaptiveTextColor.Accent;
                                if (isReminderBeforeThreeDays)
                                {
                                    reminderType = this.localizer.GetString("TeamGoalEndingAfterThreeDays");
                                    reminderTypeColor = AdaptiveTextColor.Warning;
                                }
                                else
                                {
                                    reminderType = this.GetReminderTypeString(teamGoalDetail.ReminderFrequency);
                                }

                                var goalReminderAttachment = MessageFactory.Attachment(GoalReminderCard.GetGoalReminderCard(this.localizer, this.options.Value.ManifestId, this.options.Value.GoalsTabEntityId, reminderType, reminderTypeColor));
                                this.logger.LogInformation($"Sending goal reminder card to teamId: {teamId}");
                                await turnContext.SendActivityAsync(goalReminderAttachment, cancellationToken);
                                await this.SendGoalReminderToTeamMembersAsync(turnContext, teamGoalDetail, goalReminderAttachment, cancellationToken);
                            },
                            CancellationToken.None);
                    }
                    catch (Exception ex)
                    {
                        this.logger.LogError(ex, $"Error while performing retry logic to send goal reminder card for : {teamGoalDetail.TeamGoalId}.");
                        throw;
                    }
                });
            }
            #pragma warning disable CA1031 // Catching general exceptions to log exception details in telemetry client.
            catch (Exception ex)
            #pragma warning restore CA1031 // Catching general exceptions to log exception details in telemetry client.
            {
                this.logger.LogError(ex, $"Error while sending goal reminder in team from background service for : {teamGoalDetail.TeamGoalId} at {nameof(this.SendGoalReminderToTeamAndTeamMembersAsync)}");
            }
        }

        /// <summary>
        /// Method to update team goal, personal goal and note details in storage when team goal is ended.
        /// </summary>
        /// <param name="teamGoalDetail">Holds team goal detail entity data sent from background service.</param>
        /// <returns>A task that represents team goal, personal goal and personal goal note details data is saved or updated.</returns>
        public async Task UpdateGoalDetailsAsync(TeamGoalDetail teamGoalDetail)
        {
            teamGoalDetail = teamGoalDetail ?? throw new ArgumentNullException(nameof(teamGoalDetail));

            try
            {
                var teamId = teamGoalDetail.TeamId;
                string serviceUrl = teamGoalDetail.ServiceUrl;
                MicrosoftAppCredentials.TrustServiceUrl(serviceUrl);

                var conversationReference = new ConversationReference()
                {
                    ChannelId = Constants.TeamsBotFrameworkChannelId,
                    Bot = new ChannelAccount() { Id = $"28:{this.microsoftAppCredentials.MicrosoftAppId}" },
                    ServiceUrl = serviceUrl,
                    Conversation = new ConversationAccount() { ConversationType = Constants.ChannelConversationType, IsGroup = true, Id = teamId, TenantId = this.options.Value.TenantId },
                };

                await this.retryPolicy.ExecuteAsync(async () =>
                {
                    try
                    {
                        await ((BotFrameworkAdapter)this.adapter).ContinueConversationAsync(
                            this.microsoftAppCredentials.MicrosoftAppId,
                            conversationReference,
                            async (turnContext, cancellationToken) =>
                            {
                                await this.UpdateGoalEntitiesAsync(turnContext, teamGoalDetail, cancellationToken);
                            },
                            CancellationToken.None);
                    }
                    catch (Exception ex)
                    {
                        this.logger.LogError(ex, $"Error while performing retry logic to send goal reminder card for : {teamGoalDetail.TeamGoalId}.");
                        throw;
                    }
                });
            }
            #pragma warning disable CA1031 // Catching general exceptions to log exception details in telemetry client.
            catch (Exception ex)
            #pragma warning restore CA1031 // Catching general exceptions to log exception details in telemetry client.
            {
                this.logger.LogError(ex, $"Error while updating personal goal, team goal and personal goal note detail from background service for : {teamGoalDetail.TeamGoalId} at {nameof(this.UpdateGoalDetailsAsync)}");
            }
        }

        /// <summary>
        /// Update team goal, personal goal and personal goal note details in storage if team goal cycle is ended.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="teamGoalDetail">Holds team goal detail entity data sent from background service.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents team goal, personal goal and personal goal note details data is saved or updated.</returns>
        private async Task UpdateGoalEntitiesAsync(ITurnContext turnContext, TeamGoalDetail teamGoalDetail, CancellationToken cancellationToken)
        {
            try
            {
                var teamGoalEntities = await this.teamGoalStorageProvider.GetTeamGoalDetailsByTeamIdAsync(teamGoalDetail.TeamId);

                // Update team goal details data when goal cycle is ended
                foreach (var teamGoalEntity in teamGoalEntities)
                {
                    teamGoalEntity.IsActive = false;
                    teamGoalEntity.IsReminderActive = false;
                    teamGoalEntity.LastModifiedOn = DateTime.UtcNow.ToString(CultureInfo.CurrentCulture);
                }

                await this.teamGoalStorageProvider.CreateOrUpdateTeamGoalDetailsAsync(teamGoalEntities);

                var teamMembers = await this.goalHelper.GetMembersInTeamAsync(turnContext, cancellationToken);
                foreach (var teamMember in teamMembers)
                {
                    var userAadObjectId = teamMember.AadObjectId;
                    var alignedGoalDetails = await this.personalGoalStorageProvider.GetUserAlignedGoalDetailsByTeamIdAsync(teamGoalDetail.TeamId, userAadObjectId);
                    if (alignedGoalDetails.FirstOrDefault() != null)
                    {
                        // Update personal goal details data when goal cycle is ended for aligned goals
                        foreach (var personalGoalEntity in alignedGoalDetails)
                        {
                            personalGoalEntity.IsActive = false;
                            personalGoalEntity.IsReminderActive = false;
                            personalGoalEntity.LastModifiedOn = DateTime.UtcNow.ToString(CultureInfo.CurrentCulture);

                            var personalGoalNoteEntities = await this.personalGoalNoteStorageProvider.GetPersonalGoalNoteDetailsAsync(personalGoalEntity.PersonalGoalId, userAadObjectId);

                            // Update personal goal note details data when goal cycle is ended for aligned goals
                            foreach (var personalGoalNoteEntity in personalGoalNoteEntities)
                            {
                                personalGoalNoteEntity.IsActive = false;
                                personalGoalNoteEntity.LastModifiedOn = DateTime.UtcNow.ToString(CultureInfo.CurrentCulture);
                            }

                            await this.personalGoalNoteStorageProvider.CreateOrUpdatePersonalGoalNoteDetailsAsync(personalGoalNoteEntities);
                        }

                        await this.personalGoalStorageProvider.CreateOrUpdatePersonalGoalDetailsAsync(alignedGoalDetails);
                    }
                }
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Failed to save team goal detail data to table storage at {nameof(this.UpdateGoalEntitiesAsync)} for team id: {teamGoalDetail.TeamId}");
                throw;
            }
        }

        /// <summary>
        /// Sends goal reminder card to each member of team if he/she has aligned goal with the team.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="teamGoalDetail">Team goal details obtained from storage.</param>
        /// <param name="goalReminderActivity">Goal reminder activity to send.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A Task represents goal reminder card is sent to team and team members.</returns>
        private async Task SendGoalReminderToTeamMembersAsync(ITurnContext turnContext, TeamGoalDetail teamGoalDetail, IMessageActivity goalReminderActivity, CancellationToken cancellationToken)
        {
            var teamMembers = await this.goalHelper.GetMembersInTeamAsync(turnContext, cancellationToken);
            ConversationReference conversationReference = null;

            foreach (var teamMember in teamMembers)
            {
                // Send goal reminder card to those team members who have aligned their personal goals with the team
                var alignedGoalDetails = await this.personalGoalStorageProvider.GetUserAlignedGoalDetailsByTeamIdAsync(teamGoalDetail.TeamId, teamMember.AadObjectId);
                if (alignedGoalDetails.Any())
                {
                    var conversationParameters = new ConversationParameters
                    {
                        Bot = turnContext.Activity.Recipient,
                        Members = new ChannelAccount[] { teamMember },
                        TenantId = turnContext.Activity.Conversation.TenantId,
                    };

                    try
                    {
                        await this.retryPolicy.ExecuteAsync(async () =>
                        {
                            await ((BotFrameworkAdapter)turnContext.Adapter).CreateConversationAsync(
                            teamGoalDetail.TeamId,
                            teamGoalDetail.ServiceUrl,
                            new MicrosoftAppCredentials(this.microsoftAppCredentials.MicrosoftAppId, this.microsoftAppCredentials.MicrosoftAppPassword),
                            conversationParameters,
                            async (createConversationtTurnContext, cancellationToken1) =>
                            {
                                conversationReference = createConversationtTurnContext.Activity.GetConversationReference();
                                await ((BotFrameworkAdapter)turnContext.Adapter).ContinueConversationAsync(
                                    this.microsoftAppCredentials.MicrosoftAppId,
                                    conversationReference,
                                    async (continueConversationTurnContext, continueConversationCancellationToken) =>
                                    {
                                        this.logger.LogInformation($"Sending goal reminder card to: {teamMember.Name} from team: {teamGoalDetail.TeamId}");
                                        await continueConversationTurnContext.SendActivityAsync(goalReminderActivity, continueConversationCancellationToken);
                                    },
                                    cancellationToken);
                            },
                            cancellationToken);
                        });
                    }
                    catch (Exception ex)
                    {
                        this.logger.LogError(ex, $"Error while sending goal reminder card to members of the team : {teamGoalDetail.TeamId} at {nameof(this.SendGoalReminderToTeamMembersAsync)}");
                        throw;
                    }
                }
            }
        }

        /// <summary>
        /// Method to get text based on reminder frequency.
        /// </summary>
        /// <param name="reminderFrequency">Reminder frequency i.e. Weekly/Bi-weekly/Monthly/Quarterly.</param>
        /// <returns>Return text based on reminder frequency.</returns>
        private string GetReminderTypeString(int reminderFrequency)
        {
            return reminderFrequency switch
            {
                (int)ReminderFrequency.Weekly => this.localizer.GetString("WeeklyReminderTypeString"),
                (int)ReminderFrequency.Biweekly => this.localizer.GetString("Bi-weeklyReminderTypeString"),
                (int)ReminderFrequency.Monthly => this.localizer.GetString("MonthlyReminderTypeString"),
                (int)ReminderFrequency.Quarterly => this.localizer.GetString("QuarterlyReminderTypeString"),
                _ => string.Empty,
            };
        }
    }
}
