// <copyright file="CardHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Net;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Integration.AspNet.Core;
    using Microsoft.Bot.Builder.Teams;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
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
    /// Class that handles card create/update helper methods.
    /// </summary>
    public class CardHelper
    {
        /// <summary>
        /// Represents the task module height for set personal goal, set team goal, align goal UI.
        /// </summary>
        private const int TaskModuleHeight = 600;

        /// <summary>
        /// Represents the task module width for set personal goal, set team goal, align goal UI.
        /// </summary>
        private const int TaskModuleWidth = 600;

        /// <summary>
        /// Represents the task module height for small card in case of error/warning.
        /// </summary>
        private const int TaskModuleErrorHeight = 120;

        /// <summary>
        /// Represents the task module width for medium card in case of error/warning.
        /// </summary>
        private const int TaskModuleErrorWidth = 500;

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
        /// Microsoft App credentials.
        /// </summary>
        private readonly MicrosoftAppCredentials microsoftAppCredentials;

        /// <summary>
        /// Used to run a background task using background service.
        /// </summary>
        private readonly BackgroundTaskWrapper backgroundTaskWrapper;

        /// <summary>
        /// Bot adapter.
        /// </summary>
        private readonly IBotFrameworkHttpAdapter adapter;

        /// <summary>
        /// Instance to send logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<CardHelper> logger;

        /// <summary>
        /// The current cultures' string localizer.
        /// </summary>
        private readonly IStringLocalizer<Strings> localizer;

        /// <summary>
        /// A set of key/value application configuration properties for Activity settings.
        /// </summary>
        private readonly IOptions<GoalTrackerActivityHandlerOptions> options;

        /// <summary>
        /// A set of key/value application setting related to application insights token.
        /// </summary>
        private readonly IOptions<TelemetryOptions> telemetryOptions;

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
        /// Initializes a new instance of the <see cref="CardHelper"/> class.
        /// </summary>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="options">A set of key/value application configuration properties.</param>
        /// <param name="telemetryOptions">A set of key/value application setting related to application insights token.</param>
        /// <param name="personalGoalStorageProvider">Storage provider for working with personal goal data in storage.</param>
        /// <param name="personalGoalNoteStorageProvider">Storage provider for working with personal goal note data in storage</param>
        /// <param name="teamGoalStorageProvider">Storage provider for working with team goal data in storage.</param>
        /// <param name="goalHelper">Instance of class that handles goal helper methods.</param>
        /// <param name="microsoftAppCredentials">Instance for Microsoft app credentials.</param>
        /// <param name="backgroundTaskWrapper">Instance of backgroundTaskWrapper to run a background task using IHostedService.</param>
        /// <param name="adapter">An instance of bot adapter.</param>
        public CardHelper(
            ILogger<CardHelper> logger,
            IStringLocalizer<Strings> localizer,
            IOptions<GoalTrackerActivityHandlerOptions> options,
            IOptions<TelemetryOptions> telemetryOptions,
            IPersonalGoalStorageProvider personalGoalStorageProvider,
            IPersonalGoalNoteStorageProvider personalGoalNoteStorageProvider,
            ITeamGoalStorageProvider teamGoalStorageProvider,
            GoalHelper goalHelper,
            MicrosoftAppCredentials microsoftAppCredentials,
            BackgroundTaskWrapper backgroundTaskWrapper,
            IBotFrameworkHttpAdapter adapter)
        {
            this.options = options;
            this.telemetryOptions = telemetryOptions;
            this.logger = logger;
            this.localizer = localizer;
            this.personalGoalStorageProvider = personalGoalStorageProvider;
            this.personalGoalNoteStorageProvider = personalGoalNoteStorageProvider;
            this.teamGoalStorageProvider = teamGoalStorageProvider;
            this.goalHelper = goalHelper;
            this.microsoftAppCredentials = microsoftAppCredentials;
            this.backgroundTaskWrapper = backgroundTaskWrapper;
            this.adapter = adapter;
        }

        /// <summary>
        /// Convert Date time format to adaptive card text feature.
        /// </summary>
        /// <param name="inputText">Input date time string.</param>
        /// <returns>Adaptive card supported date time format.</returns>
        public static string FormatDateStringToAdaptiveCardDateFormat(string inputText)
        {
            try
            {
                return "{{DATE(" + DateTime.Parse(inputText, CultureInfo.InvariantCulture).ToUniversalTime().ToString(Constants.Rfc3339DateTimeFormat, CultureInfo.InvariantCulture) + ", SHORT)}}";
            }
#pragma warning disable CA1031 // Do not catch general exception types
            catch
#pragma warning restore CA1031 // Do not catch general exception types
            {
                return inputText;
            }
        }

        /// <summary>
        /// Get add note card task module invoked through bot command or button click.
        /// </summary>
        /// <param name="personalGoalNoteDetail">Holds personal goal note detail entity data.</param>
        /// <param name="personalGoalDetail">Holds collection of personal goal detail entity data.</param>
        /// <param name="addNoteTaskModuleTitle">Add note title text to be shown in task module.</param>
        /// <returns>Returns add note card.</returns>
        public TaskModuleResponse GetAddNoteCardResponse(PersonalGoalNoteDetail personalGoalNoteDetail, IEnumerable<PersonalGoalDetail> personalGoalDetail, string addNoteTaskModuleTitle)
        {
            return new TaskModuleResponse
            {
                Task = new TaskModuleContinueResponse
                {
                    Value = new TaskModuleTaskInfo()
                    {
                        Card = NoteCard.GetAddNoteCardInTaskModule(personalGoalNoteDetail, personalGoalDetail, this.localizer),
                        Height = TaskModuleHeight,
                        Width = TaskModuleWidth,
                        Title = addNoteTaskModuleTitle,
                    },
                },
            };
        }

        /// <summary>
        /// Get add note card task module invoked through message action.
        /// </summary>
        /// <param name="personalGoalNoteDetail">Holds personal goal note detail entity data.</param>
        /// <param name="personalGoalDetail">Holds collection of personal goal detail entity data.</param>
        /// <returns>Returns add note card.</returns>
        public async Task<MessagingExtensionActionResponse> GetAddNoteCardResponseAsync(PersonalGoalNoteDetail personalGoalNoteDetail, IEnumerable<PersonalGoalDetail> personalGoalDetail)
        {
            return await Task.FromResult(new MessagingExtensionActionResponse
            {
                Task = new TaskModuleContinueResponse
                {
                    Value = new TaskModuleTaskInfo
                    {
                        Card = NoteCard.GetAddNoteCardInTaskModule(personalGoalNoteDetail, personalGoalDetail, this.localizer),
                        Height = TaskModuleHeight,
                        Width = TaskModuleWidth,
                        Title = this.localizer.GetString("AddNoteButtonText"),
                    },
                },
            });
        }

        /// <summary>
        /// Get goal cycle validation response for personal goal and team goal for fetch action.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="goalCycleId">Goal cycle id to identify that goal cycle is ended or not.</param>
        /// <param name="taskModuleErrorMessage">Error message to be shown in task module.</param>
        /// <param name="taskModuleTitle">Task module title for error response task module.</param>
        /// <param name="teamId">Team id to fetch team goal detail.</param>
        /// <param name="isPersonalGoalCycleValidation">Determines whether validation is for personal goal cycle.</param>
        /// <returns>Returns goal cycle validation card for fetch action.</returns>
        public async Task<TaskModuleResponse> GetGoalCycleValidationResponseAsync(ITurnContext<IInvokeActivity> turnContext, string goalCycleId, string taskModuleErrorMessage, string taskModuleTitle, string teamId = null, bool isPersonalGoalCycleValidation = false)
        {
            turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
            if (isPersonalGoalCycleValidation)
            {
                var personalGoalDetail = (await this.personalGoalStorageProvider.GetPersonalGoalDetailsByUserAadObjectIdAsync(turnContext.Activity.From.AadObjectId)).FirstOrDefault();
                if (personalGoalDetail == null || (!string.IsNullOrEmpty(goalCycleId) && personalGoalDetail.GoalCycleId != goalCycleId))
                {
                    this.logger.LogError($"The goals, user: {turnContext.Activity.From.Name} is looking for has been deleted from storage or goal cycle has been ended.");
                    return this.GetTaskModuleErrorResponse(taskModuleErrorMessage, taskModuleTitle);
                }

                return null;
            }

            var teamGoalDetail = (await this.teamGoalStorageProvider.GetTeamGoalDetailsByTeamIdAsync(teamId)).FirstOrDefault();
            if (teamGoalDetail == null || (!string.IsNullOrEmpty(goalCycleId) && teamGoalDetail.GoalCycleId != goalCycleId))
            {
                this.logger.LogError($"The goals for teamId: {teamId} has been deleted from storage or goal cycle has been ended");
                return this.GetTaskModuleErrorResponse(this.localizer.GetString("NoActiveGoalOrGoalCycleEndedError"), this.localizer.GetString("EditTeamGoalTaskModuleTitle"));
            }

            return null;
        }

        /// <summary>
        /// Get validation card response for add note for submit action.
        /// </summary>
        /// <param name="personalGoalNoteDetail">Holds personal goal note detail entity data.</param>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <returns>Returns add note validation card for submit action.</returns>
        public async Task<TaskModuleResponse> GetAddNoteErrorResponseAsync(PersonalGoalNoteDetail personalGoalNoteDetail, ITurnContext<IInvokeActivity> turnContext)
        {
            turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
            if (personalGoalNoteDetail == null)
            {
                this.logger.LogInformation("Value obtained for personal goal note detail from task module submit action is null");
                return this.GetTaskModuleErrorResponse(this.localizer.GetString("GenericErrorMessage"), this.localizer.GetString("AddNoteTaskModuleTitle"));
            }

            var personalGoalDetails = await this.personalGoalStorageProvider.GetPersonalGoalDetailsByUserAadObjectIdAsync(turnContext.Activity.From.AadObjectId);
            if (string.IsNullOrWhiteSpace(personalGoalNoteDetail.PersonalGoalNoteDescription) || string.IsNullOrWhiteSpace(personalGoalNoteDetail.PersonalGoalId))
            {
                this.logger.LogInformation("Show validation message in add note card task module on submit action");
                return this.GetAddNoteValidationCardResponse(personalGoalNoteDetail, personalGoalDetails);
            }

            int numberOfNotes = personalGoalNoteDetail.PersonalGoalId == null ? 0 : await this.personalGoalNoteStorageProvider.GetNumberOfNotesForGoalAsync(personalGoalNoteDetail.PersonalGoalId, turnContext.Activity.From.AadObjectId);
            if (numberOfNotes > Constants.MaximumNumberOfNotes)
            {
                this.logger.LogInformation("Show validation message in add note card task module on submit action");
                return this.GetAddNoteValidationCardResponse(personalGoalNoteDetail, personalGoalDetails, numberOfNotes);
            }

            return null;
        }

        /// <summary>
        /// Get validation card response for add note.
        /// </summary>
        /// <param name="personalGoalNoteDetail">Holds personal goal note detail entity data.</param>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <returns>Returns add note validation card.</returns>
        public async Task<MessagingExtensionActionResponse> GetAddNoteErrorResponseForMessageActionAsync(PersonalGoalNoteDetail personalGoalNoteDetail, ITurnContext<IInvokeActivity> turnContext)
        {
            turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
            if (personalGoalNoteDetail == null)
            {
                this.logger.LogInformation("Value obtained for personal goal note detail from task module submit using message action is null");
                return await this.GetTaskModuleErrorResponseAsync(this.localizer.GetString("GenericErrorMessage"), this.localizer.GetString("AddNoteTaskModuleTitle"));
            }

            var personalGoalDetails = await this.personalGoalStorageProvider.GetPersonalGoalDetailsByUserAadObjectIdAsync(turnContext.Activity.From.AadObjectId);
            if (string.IsNullOrWhiteSpace(personalGoalNoteDetail.PersonalGoalNoteDescription) || string.IsNullOrWhiteSpace(personalGoalNoteDetail.PersonalGoalId))
            {
                this.logger.LogInformation("Show validation message in add note card task module on submit action invoked through message action");
                return this.GetAddNoteValidationCardResponseForMessageAction(personalGoalNoteDetail, personalGoalDetails);
            }

            int numberOfNotes = personalGoalNoteDetail.PersonalGoalId == null ? 0 : await this.personalGoalNoteStorageProvider.GetNumberOfNotesForGoalAsync(personalGoalNoteDetail.PersonalGoalId, turnContext.Activity.From.AadObjectId);
            if (numberOfNotes > Constants.MaximumNumberOfNotes)
            {
                this.logger.LogInformation("Show validation message in add note card task module on submit action invoked through message action");
                return this.GetAddNoteValidationCardResponseForMessageAction(personalGoalNoteDetail, personalGoalDetails, numberOfNotes);
            }

            turnContext.Activity.Conversation.Id = personalGoalDetails.First().ConversationId;
            return null;
        }

        /// <summary>
        /// Get validation card response for add note for fetch action.
        /// </summary>
        /// <param name="personalGoalDetails">Holds collection of personal goal detail entity data.</param>
        /// <param name="goalCycleId">Goal cycle id to identify that goal cycle is ended or not.</param>
        /// <returns>Returns add note validation card for fetch action.</returns>
        public TaskModuleResponse GetPersonalNoteValidationResponse(IEnumerable<PersonalGoalDetail> personalGoalDetails, string goalCycleId)
        {
            if (!personalGoalDetails.Any())
            {
                this.logger.LogError("No personal goals present for adding note");
                return this.GetTaskModuleErrorResponse(this.localizer.GetString("AddNoteNoPersonalGoalAvailable"), this.localizer.GetString("AddNoteTaskModuleTitle"));
            }

            if (!string.IsNullOrEmpty(goalCycleId) && personalGoalDetails.First().GoalCycleId != goalCycleId)
            {
                this.logger.LogInformation($"Goal cycle has been ended.");
                return this.GetTaskModuleErrorResponse(this.localizer.GetString("GoalCycleEndedErrorForAddingNote"), this.localizer.GetString("AddNoteTaskModuleTitle"));
            }

            return null;
        }

        /// <summary>
        /// Get validation card response for add note.
        /// </summary>
        /// <param name="personalGoalNoteDetail">Holds personal goal note detail entity data.</param>
        /// <param name="personalGoalDetails">Holds collection of personal goal detail entity data.</param>
        /// <param name="numberOfNotes">Number of goal notes for a particular goal.</param>
        /// <returns>Returns add note validation card.</returns>
        public TaskModuleResponse GetAddNoteValidationCardResponse(PersonalGoalNoteDetail personalGoalNoteDetail, IEnumerable<PersonalGoalDetail> personalGoalDetails, int numberOfNotes = 0)
        {
            return new TaskModuleResponse
            {
                Task = new TaskModuleContinueResponse
                {
                    Value = new TaskModuleTaskInfo()
                    {
                        Card = NoteCard.GetAddNoteCardInTaskModule(personalGoalNoteDetail, personalGoalDetails, this.localizer, string.IsNullOrWhiteSpace(personalGoalNoteDetail?.PersonalGoalNoteDescription), string.IsNullOrWhiteSpace(personalGoalNoteDetail?.PersonalGoalId), numberOfNotes > Constants.MaximumNumberOfNotes),
                        Height = TaskModuleHeight,
                        Width = TaskModuleWidth,
                        Title = this.localizer.GetString("AddNoteTaskModuleTitle"),
                    },
                },
            };
        }

        /// <summary>
        /// Get add note validation card in message action.
        /// </summary>
        /// <param name="personalGoalNoteDetail">Holds personal goal note detail entity data.</param>
        /// <param name="personalGoalDetail">Holds collection of personal goal detail entity data.</param>
        /// <param name="numberOfGoalNotes">Number of goal Notes for a particular goal</param>
        /// <returns>Returns add note validation card.</returns>
        public MessagingExtensionActionResponse GetAddNoteValidationCardResponseForMessageAction(PersonalGoalNoteDetail personalGoalNoteDetail, IEnumerable<PersonalGoalDetail> personalGoalDetail, int numberOfGoalNotes = 0)
        {
            return new MessagingExtensionActionResponse
            {
                Task = new TaskModuleContinueResponse
                {
                    Value = new TaskModuleTaskInfo()
                    {
                        Card = NoteCard.GetAddNoteCardInTaskModule(personalGoalNoteDetail, personalGoalDetail, this.localizer, string.IsNullOrWhiteSpace(personalGoalNoteDetail?.PersonalGoalNoteDescription), string.IsNullOrWhiteSpace(personalGoalNoteDetail?.PersonalGoalId), numberOfGoalNotes > Constants.MaximumNumberOfNotes),
                        Height = TaskModuleHeight,
                        Width = TaskModuleWidth,
                        Title = this.localizer.GetString("UpdateStatusTitle"),
                    },
                },
            };
        }

        /// <summary>
        /// Get set goal adaptive card.
        /// </summary>
        /// <param name="activityServicePath">Activity service URL</param>
        /// <param name="setPersonalGoalTaskModuleTitle">Set/edit personal goal task module title.</param>
        /// <returns>Returns set goal card to be displayed in task module.</returns>
        public TaskModuleResponse GetPersonalSetGoalCardResponse(string activityServicePath, string setPersonalGoalTaskModuleTitle)
        {
            return new TaskModuleResponse
            {
                Task = new TaskModuleContinueResponse
                {
                    Value = new TaskModuleTaskInfo()
                    {
                        Url = $"{this.options.Value.AppBaseUri}/personal-goal?telemetry={this.telemetryOptions.Value.InstrumentationKey}&serviceURL={activityServicePath}",
                        Height = TaskModuleHeight,
                        Width = TaskModuleWidth,
                        Title = setPersonalGoalTaskModuleTitle,
                    },
                },
            };
        }

        /// <summary>
        /// Store or update personal goal note detail in storage and sends/updates note summary card.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="personalGoalNoteDetail">Holds collection of personal goal note detail entity data.</param>
        /// <param name="adaptiveSubmitActionData">Holds collection of submit action data.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>A task represents personal goal note details are saved in the storage successfully.</returns>
        public async Task StoreOrUpdateNoteDetailAsync(ITurnContext<IInvokeActivity> turnContext, PersonalGoalNoteDetail personalGoalNoteDetail, AdaptiveSubmitActionData adaptiveSubmitActionData, CancellationToken cancellationToken)
        {
            turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
            personalGoalNoteDetail = personalGoalNoteDetail ?? throw new ArgumentNullException(nameof(personalGoalNoteDetail));
            adaptiveSubmitActionData = adaptiveSubmitActionData ?? throw new ArgumentNullException(nameof(adaptiveSubmitActionData));

            var activity = turnContext.Activity;
            var personalGoalDetailByGoalId = await this.personalGoalStorageProvider.GetPersonalGoalDetailByGoalIdAsync(personalGoalNoteDetail.PersonalGoalId, activity.From.AadObjectId);

            if (string.IsNullOrEmpty(adaptiveSubmitActionData.GoalNoteId))
            {
                personalGoalNoteDetail.UserAadObjectId = activity.From.AadObjectId;
                personalGoalNoteDetail.ConversationId = activity.Conversation.Id;
                personalGoalNoteDetail.PersonalGoalNoteId = Guid.NewGuid().ToString();
                personalGoalNoteDetail.IsActive = true;
                personalGoalNoteDetail.CreatedOn = personalGoalDetailByGoalId.CreatedOn;
                personalGoalNoteDetail.CreatedBy = personalGoalDetailByGoalId.CreatedBy;
                personalGoalNoteDetail.LastModifiedOn = DateTime.UtcNow.ToString(Constants.Rfc3339DateTimeFormat, CultureInfo.CurrentCulture);
                personalGoalNoteDetail.LastModifiedBy = activity.From.Name;
                var savePersonalGoalNoteData = await this.personalGoalNoteStorageProvider.CreateOrUpdatePersonalGoalNoteDetailAsync(personalGoalNoteDetail);
                if (!savePersonalGoalNoteData)
                {
                    this.logger.LogInformation("Error while saving personal goal notes to storage");
                    await turnContext.SendActivityAsync(this.localizer.GetString("GenericErrorMessage"), cancellationToken: cancellationToken);
                    return;
                }

                var addNoteCardAttachment = MessageFactory.Attachment(NoteCard.GetsAddNoteSubmitCard(personalGoalNoteDetail, personalGoalDetailByGoalId, this.localizer));
                this.logger.LogInformation("Sending add note card to personal bot");
                var addNoteSubmitCardActivity = await turnContext.SendActivityAsync(addNoteCardAttachment, cancellationToken);

                // Update activity response id of the card in storage for updating the conversation later
                personalGoalNoteDetail.AdaptiveCardActivityId = addNoteSubmitCardActivity.Id;
                savePersonalGoalNoteData = await this.personalGoalNoteStorageProvider.CreateOrUpdatePersonalGoalNoteDetailAsync(personalGoalNoteDetail);
                if (!savePersonalGoalNoteData)
                {
                    this.logger.LogInformation("Error while saving personal goal notes to storage");
                    await turnContext.SendActivityAsync(this.localizer.GetString("GenericErrorMessage"), cancellationToken: cancellationToken);
                    return;
                }
            }
            else
            {
                var savedPersonalGoalNoteDetail = await this.personalGoalNoteStorageProvider.GetPersonalGoalNoteDetailAsync(adaptiveSubmitActionData.GoalNoteId, activity.From.AadObjectId);
                savedPersonalGoalNoteDetail.UserAadObjectId = turnContext.Activity.From.AadObjectId;
                savedPersonalGoalNoteDetail.PersonalGoalId = personalGoalNoteDetail.PersonalGoalId;
                savedPersonalGoalNoteDetail.PersonalGoalNoteDescription = personalGoalNoteDetail.PersonalGoalNoteDescription;
                savedPersonalGoalNoteDetail.SourceName = personalGoalNoteDetail.SourceName;
                savedPersonalGoalNoteDetail.LastModifiedOn = DateTime.UtcNow.ToString(Constants.Rfc3339DateTimeFormat, CultureInfo.CurrentCulture);
                savedPersonalGoalNoteDetail.LastModifiedBy = turnContext.Activity.From.Name;

                var updatedAddNoteCardAttachment = MessageFactory.Attachment(NoteCard.GetsAddNoteSubmitCard(savedPersonalGoalNoteDetail, personalGoalDetailByGoalId, this.localizer));
                updatedAddNoteCardAttachment.Id = savedPersonalGoalNoteDetail.AdaptiveCardActivityId;
                updatedAddNoteCardAttachment.Conversation = turnContext.Activity.Conversation;
                this.logger.LogInformation("Updating note card in personal bot");
                await turnContext.UpdateActivityAsync(updatedAddNoteCardAttachment, cancellationToken);
                var updatePersonalGoalNoteData = await this.personalGoalNoteStorageProvider.CreateOrUpdatePersonalGoalNoteDetailAsync(savedPersonalGoalNoteDetail);
                if (!updatePersonalGoalNoteData)
                {
                    this.logger.LogInformation("Error while saving personal goal notes to storage");
                    await turnContext.SendActivityAsync(this.localizer.GetString("GenericErrorMessage"), cancellationToken: cancellationToken);
                    return;
                }
            }
        }

        /// <summary>
        /// Get valid aligned goal details data from storage.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="teamId">Team id to get teams goal details from storage.</param>
        /// <returns>Return valid aligned goal details.</returns>
        public async Task<TaskModuleResponse> GetAlignGoalTaskModuleResponseAsync(ITurnContext<IInvokeActivity> turnContext, string teamId)
        {
            try
            {
                turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
                var personalGoalDetails = await this.personalGoalStorageProvider.GetPersonalGoalDetailsByUserAadObjectIdAsync(turnContext.Activity.From.AadObjectId);
                var personalGoalDetail = personalGoalDetails.FirstOrDefault();
                if (personalGoalDetail == null)
                {
                    this.logger.LogInformation("Fetch task module to show error card when personal goal details are null");
                    return this.GetTaskModuleErrorResponse(string.Format(CultureInfo.InvariantCulture, this.localizer.GetString("AlignGoalNoPersonalGoalsAdded"), turnContext.Activity.From.Name), this.localizer.GetString("AlignGoalTaskModuleTitle"));
                }

                var teamGoalDetail = (await this.teamGoalStorageProvider.GetTeamGoalDetailsByTeamIdAsync(teamId)).FirstOrDefault();
                if (teamGoalDetail == null)
                {
                    this.logger.LogInformation("Fetch task module to show error card when team goal note details are null");
                    return this.GetTaskModuleErrorResponse(string.Format(CultureInfo.InvariantCulture, this.localizer.GetString("AlignGoalNoTeamGoalsAdded"), turnContext.Activity.From.Name), this.localizer.GetString("AlignGoalTaskModuleTitle"));
                }

                var personalGoalStartDate = DateTime.Parse(personalGoalDetail.StartDate, CultureInfo.CurrentCulture).ToUniversalTime();
                var personalGoalEndDate = DateTime.Parse(personalGoalDetail.EndDate, CultureInfo.CurrentCulture).ToUniversalTime();
                var teamGoalStartDate = DateTime.Parse(teamGoalDetail.TeamGoalStartDate, CultureInfo.CurrentCulture).ToUniversalTime();
                var teamGoalEndDate = DateTime.Parse(teamGoalDetail.TeamGoalEndDate, CultureInfo.CurrentCulture).ToUniversalTime();
                if ((teamGoalStartDate > personalGoalStartDate || personalGoalStartDate > teamGoalEndDate)
                   && (teamGoalStartDate > personalGoalEndDate || personalGoalEndDate > teamGoalEndDate)
                   && (personalGoalStartDate > teamGoalStartDate || personalGoalEndDate < teamGoalStartDate)
                   && (teamGoalStartDate > personalGoalStartDate || teamGoalEndDate < personalGoalEndDate))
                {
                    this.logger.LogInformation("Fetch task module to show error card when personal goal start date and end date does not fall on or between team goal start date and end date");
                    return this.GetTaskModuleErrorResponse(string.Format(CultureInfo.InvariantCulture, this.localizer.GetString("AlignGoalDateConflict")), this.localizer.GetString("AlignGoalTaskModuleTitle"));
                }

                var alignedTeamDetail = personalGoalDetails.Where(personalGoalDetail1 => !string.IsNullOrEmpty(personalGoalDetail1.TeamId) && personalGoalDetail1.IsAligned).FirstOrDefault();
                if (alignedTeamDetail != null && alignedTeamDetail.TeamId != teamId)
                {
                    this.logger.LogInformation("Fetch task module to show alignment change confirmation card");
                    return this.GetAlignmentChangeConfirmationCardResponse(teamId);
                }
                else
                {
                    this.logger.LogInformation("Fetch task module to show align goal card");
                    return this.GetAlignGoalCardResponse(teamId);
                }
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"An error occurred in {nameof(this.GetAlignGoalTaskModuleResponseAsync)} while fetching goal alignment details for : {teamId}");
                throw;
            }
        }

        /// <summary>
        /// Gets team member details and sends list card to members of the team.
        /// </summary>
        /// <param name="applicationBasePath">Application base URL of the application.</param>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="teamGoalDetail">Team goal values entered by user.</param>
        /// <param name="goalCycleId">Goal cycle id to identify that goal cycle is ended or not.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>Returns the task that sends list card to all members of the team.</returns>
        public async Task SendTeamGoalListCardToTeamMembersAsync(string applicationBasePath, ITurnContext<IInvokeActivity> turnContext, IEnumerable<TeamGoalDetail> teamGoalDetail, string goalCycleId, CancellationToken cancellationToken)
        {
            try
            {
                turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
                var teamMembers = await this.goalHelper.GetMembersInTeamAsync(turnContext, cancellationToken);
                if (teamMembers == null)
                {
                    this.logger.LogError("No team members found.");
                    return;
                }

                var teamGoalListCardAttachmentForTeamMembers = MessageFactory.Attachment(GoalCard.GetTeamGoalDetailsListCard(applicationBasePath, teamGoalDetail, this.localizer, turnContext.Activity.From.Name, goalCycleId, true));
                string serviceURL = teamGoalDetail.ToList().Select(teamGoal => teamGoal.ServiceUrl).FirstOrDefault();

                // Send list card to all members of the team. Enqueue task to task wrapper and it will be executed by goal background service.
                this.backgroundTaskWrapper.Enqueue(this.SendListCardToMembersAsync(turnContext, teamMembers, teamGoalListCardAttachmentForTeamMembers, serviceURL, cancellationToken));
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"An error occurred in {nameof(this.SendTeamGoalListCardToTeamMembersAsync)}");
                throw;
            }
        }

        /// <summary>
        /// Method to update personal goal list card when personal goal detail updated/deleted from personal tab.
        /// </summary>
        /// <param name="personalGoalDetail">Personal goal detail entity received from client application.</param>
        /// <returns>A task represent personal goal list card of personal bot to be updated.</returns>
        public async Task UpdatePersonalGoalListCardAsync(PersonalGoalDetail personalGoalDetail)
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
                            var personalGoalDetails = await this.personalGoalStorageProvider.GetPersonalGoalDetailsByUserAadObjectIdAsync(personalGoalDetail.UserAadObjectId);
                            var goalCyleId = personalGoalDetails.First().GoalCycleId;
                            var personalGoalListCardAttacment = MessageFactory.Attachment(GoalCard.GetPersonalGoalDetailsListCard(this.options.Value.AppBaseUri, personalGoalDetails, this.localizer, goalCyleId));

                            // Update personal goals list card in personal chat.
                            personalGoalListCardAttacment.Id = personalGoalDetail.AdaptiveCardActivityId;
                            this.logger.LogInformation("Updating personal goal list card");
                            await turnContext.UpdateActivityAsync(personalGoalListCardAttacment, cancellationToken);
                        },
                        CancellationToken.None);
                });
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error while updating personal goal list card for personal bot: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Method to update personal goal note card when personal goal note details updated/deleted from personal tab.
        /// </summary>
        /// <param name="personalGoalNoteDetails">Personal goal note details entities received from client application.</param>
        /// <returns>A task represent personal goal note card of personal bot to be updated.</returns>
        public async Task UpdatePersonalNoteCardAsync(IEnumerable<PersonalGoalNoteDetail> personalGoalNoteDetails)
        {
            personalGoalNoteDetails = personalGoalNoteDetails ?? throw new ArgumentNullException(nameof(personalGoalNoteDetails));

            var personalGoalId = personalGoalNoteDetails.First().PersonalGoalId;
            var userAadObjectId = personalGoalNoteDetails.First().UserAadObjectId;
            var personalGoalDetail = await this.personalGoalStorageProvider.GetPersonalGoalDetailByGoalIdAsync(personalGoalId, userAadObjectId);

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
                            foreach (var personalGoalNoteDetail in personalGoalNoteDetails)
                            {
                                var updatedAddNoteCardAttachment = MessageFactory.Attachment(NoteCard.GetsAddNoteSubmitCard(personalGoalNoteDetail, personalGoalDetail, this.localizer));
                                updatedAddNoteCardAttachment.Id = personalGoalNoteDetail.AdaptiveCardActivityId;
                                this.logger.LogInformation("Updating note card in personal bot");
                                await turnContext.UpdateActivityAsync(updatedAddNoteCardAttachment, cancellationToken);
                            }
                        },
                        CancellationToken.None);
                });
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error while updating note card for personal bot: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Get valid aligned goal details data from storage.
        /// </summary>
        /// <param name="userAadObjectId">AAD object id of the user whose goal details needs to be deleted.</param>
        /// <param name="teamId">Team id to get teams goal details from storage.</param>
        /// <returns>Return valid aligned goal details.</returns>
        public async Task<TaskModuleResponse> GetAlignGoalTaskModuleResponseAsync(string userAadObjectId, string teamId)
        {
            try
            {
                if (string.IsNullOrEmpty(teamId))
                {
                    // When user clicks 'No' on alignment change confirmation dialog, task module will be closed.
                    return null;
                }

                // Unaligned team goals from personal goal details when user want to align goals in another team:
                var personalGoalDetails = await this.personalGoalStorageProvider.GetPersonalGoalDetailsByUserAadObjectIdAsync(userAadObjectId);
                personalGoalDetails = personalGoalDetails.Where(personalGoalDetail => personalGoalDetail.IsAligned).ToList();
                var deleteAlignedGoals = await this.goalHelper.DeleteAlignedGoalDetailsAsync(personalGoalDetails);
                if (!deleteAlignedGoals)
                {
                    this.logger.LogInformation("Error while deleting aligned goal details from storage.");
                    return this.GetTaskModuleErrorResponse(this.localizer.GetString("AlignGoalErrorDeletingData"), this.localizer.GetString("AlignGoalTaskModuleTitle"));
                }

                this.logger.LogInformation("Fetch task module to show align goal card");
                return this.GetAlignGoalCardResponse(teamId);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"An error occurred in {nameof(this.GetAlignGoalTaskModuleResponseAsync)} while fetching goal alignment details for: {teamId}");
                throw;
            }
        }

        /// <summary>
        /// Get align goal adaptive card in response to align goal in different channels.
        /// </summary>
        /// <param name="teamId">Team id to align personal goal within specified team.</param>
        /// <returns>Returns align goal card to be displayed in task module.</returns>
        public TaskModuleResponse GetAlignGoalCardResponse(string teamId)
        {
            return new TaskModuleResponse
            {
                Task = new TaskModuleContinueResponse
                {
                    Value = new TaskModuleTaskInfo()
                    {
                        Url = $"{this.options.Value.AppBaseUri}/align-goal?teamId={teamId}&telemetry={this.telemetryOptions.Value.InstrumentationKey}",
                        Height = TaskModuleHeight,
                        Width = TaskModuleWidth,
                        Title = this.localizer.GetString("AlignGoalTaskModuleTitle"),
                    },
                },
            };
        }

        /// <summary>
        /// Gets the confirmation card in response for changing goal alignment from one team to another.
        /// </summary>
        /// <param name="teamId">Team id to align personal goal within specified team.</param>
        /// <returns>Returns confirmation goal alignment change card to be displayed in task module.</returns>
        public TaskModuleResponse GetAlignmentChangeConfirmationCardResponse(string teamId)
        {
            return new TaskModuleResponse
            {
                Task = new TaskModuleContinueResponse
                {
                    Value = new TaskModuleTaskInfo()
                    {
                        Card = GoalCard.GetAlignmentChangeConfirmationCard(this.localizer, teamId),
                        Height = TaskModuleErrorHeight,
                        Width = TaskModuleErrorWidth,
                        Title = this.localizer.GetString("AlignGoalTaskModuleTitle"),
                    },
                },
            };
        }

        /// <summary>
        /// Get the error card task module on validation failure.
        /// </summary>
        /// <param name="errorMessage">Error message to be displayed in task module.</param>
        /// <param name="title">Title for task module describing type of error.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        public TaskModuleResponse GetTaskModuleErrorResponse(string errorMessage, string title)
        {
            return new TaskModuleResponse
            {
                Task = new TaskModuleContinueResponse
                {
                    Value = new TaskModuleTaskInfo()
                    {
                        Card = ErrorCard.GetErrorCardAttachment(errorMessage),
                        Height = TaskModuleErrorHeight,
                        Width = TaskModuleErrorWidth,
                        Title = title,
                    },
                },
            };
        }

        /// <summary>
        /// Get the error card task module for specified error message invoked through message action.
        /// </summary>
        /// <param name="errorMessage">Error message to be displayed in task module.</param>
        /// <param name="title">Title for task module describing type of error.</param>
        /// <returns>Returns error card.</returns>
        public async Task<MessagingExtensionActionResponse> GetTaskModuleErrorResponseAsync(string errorMessage, string title)
        {
            return await Task.FromResult(new MessagingExtensionActionResponse
            {
                Task = new TaskModuleContinueResponse
                {
                    Value = new TaskModuleTaskInfo
                    {
                        Card = ErrorCard.GetErrorCardAttachment(errorMessage),
                        Height = TaskModuleErrorHeight,
                        Width = TaskModuleErrorWidth,
                        Title = title,
                    },
                },
            });
        }

        /// <summary>
        /// Get set goal adaptive card for teams.
        /// </summary>
        /// <param name="activityServicePath">ServiceURL for the team context.</param>
        /// <param name="setTeamGoalTaskModuleTitle">Set/edit team goal task module title.</param>
        /// <returns>Returns set goal card to be displayed in task module.</returns>
        public TaskModuleResponse GetTeamSetGoalCardResponse(string activityServicePath, string setTeamGoalTaskModuleTitle)
        {
            return new TaskModuleResponse
            {
                Task = new TaskModuleContinueResponse
                {
                    Value = new TaskModuleTaskInfo()
                    {
                        Url = $"{this.options.Value.AppBaseUri}/team-goal?telemetry={this.telemetryOptions.Value.InstrumentationKey}&serviceURL={activityServicePath}",
                        Height = TaskModuleHeight,
                        Width = TaskModuleWidth,
                        Title = setTeamGoalTaskModuleTitle,
                    },
                },
            };
        }

        /// <summary>
        /// Method to send list card to members of the team.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="teamMembers">Team members of the team.</param>
        /// <param name="listCard">List card containing the goals.</param>
        /// <param name="serviceURL">Service URL of team context to send card in chat.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>Returns a task to send list card to the team members.</returns>
        private async Task SendListCardToMembersAsync(ITurnContext<IInvokeActivity> turnContext, IEnumerable<TeamsChannelAccount> teamMembers, IMessageActivity listCard, string serviceURL, CancellationToken cancellationToken)
        {
            turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
            var teamsChannelId = turnContext.Activity.TeamsGetChannelId();
            var serviceUrl = serviceURL;
            var credentials = new MicrosoftAppCredentials(this.microsoftAppCredentials.MicrosoftAppId, this.microsoftAppCredentials.MicrosoftAppPassword);
            ConversationReference conversationReference = null;

            foreach (var teamMember in teamMembers)
            {
                var conversationParameters = new ConversationParameters
                {
                    IsGroup = false,
                    Bot = turnContext.Activity.Recipient,
                    Members = new ChannelAccount[] { teamMember },
                    TenantId = turnContext.Activity.Conversation.TenantId,
                };

                try
                {
                    await this.retryPolicy.ExecuteAsync(async () =>
                    {
                        await ((BotFrameworkAdapter)turnContext.Adapter).CreateConversationAsync(
                            teamsChannelId,
                            serviceUrl,
                            credentials,
                            conversationParameters,
                            async (conversationTurnContext, conversationCancellationToken) =>
                            {
                                conversationReference = conversationTurnContext.Activity.GetConversationReference();
                                await ((BotFrameworkAdapter)turnContext.Adapter).ContinueConversationAsync(
                                    this.microsoftAppCredentials.MicrosoftAppId,
                                    conversationReference,
                                    async (conversationContext, conversationCancellation) =>
                                    {
                                        this.logger.LogInformation($"Sending team goal list card to team member: {teamMember.UserPrincipalName}");
                                        await conversationContext.SendActivityAsync(listCard, conversationCancellation);
                                    },
                                    cancellationToken);
                            }, cancellationToken);
                    });
                }
                catch (Exception ex)
                {
                    this.logger.LogError(ex, $"Error while sending list card to members of the channel with id : {teamsChannelId}.");
                    throw;
                }
            }
        }
    }
}