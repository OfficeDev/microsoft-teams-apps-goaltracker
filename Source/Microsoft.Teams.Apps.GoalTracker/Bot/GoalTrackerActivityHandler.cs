// <copyright file="GoalTrackerActivityHandler.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights;
    using Microsoft.ApplicationInsights.DataContracts;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Teams;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Localization;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.GoalTracker.Cards;
    using Microsoft.Teams.Apps.GoalTracker.Common;
    using Microsoft.Teams.Apps.GoalTracker.Helpers;
    using Microsoft.Teams.Apps.GoalTracker.Models;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// The GoalTrackerActivityHandler is responsible for reacting to incoming events from Teams sent from BotFramework.
    /// </summary>
    public sealed class GoalTrackerActivityHandler : TeamsActivityHandler
    {
        /// <summary>
        /// Represents string for content type 'html' to identify content type of message action.
        /// </summary>
        private const string ContentTypeHtml = "html";

        /// <summary>
        /// Instance to send logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<GoalTrackerActivityHandler> logger;

        /// <summary>
        /// The current cultures' string localizer.
        /// </summary>
        private readonly IStringLocalizer<Strings> localizer;

        /// <summary>
        /// Instance of Application Insights Telemetry client.
        /// </summary>
        private readonly TelemetryClient telemetryClient;

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
        /// Provider for fetching information about team details from storage table.
        /// </summary>
        private readonly ITeamStorageProvider teamStorageProvider;

        /// <summary>
        /// Instance of class that handles card create/update helper methods.
        /// </summary>
        private readonly CardHelper cardHelper;

        /// <summary>
        /// Instance of class that handles goal create/update helper methods.
        /// </summary>
        private readonly GoalHelper goalHelper;

        /// <summary>
        /// Instance of class that handles Bot activity helper methods.
        /// </summary>
        private readonly ActivityHelper activityHelper;

        /// <summary>
        /// Instance of search service for working with personal goal data in storage.
        /// </summary>
        private readonly IPersonalGoalSearchService personalGoalSearchService;

        /// <summary>
        /// State management object for maintaining user conversation state.
        /// </summary>
        private readonly BotState userState;

        /// <summary>
        /// Initializes a new instance of the <see cref="GoalTrackerActivityHandler"/> class.
        /// </summary>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="telemetryClient">The Application Insights telemetry client. </param>
        /// <param name="options">A set of key/value application configuration properties.</param>
        /// <param name="personalGoalStorageProvider">Storage provider for working with personal goal data in storage.</param>
        /// <param name="personalGoalNoteStorageProvider">Storage provider for working with personal goal note data in storage</param>
        /// <param name="teamGoalStorageProvider">Storage provider for working with team goal data in storage.</param>
        /// <param name="teamStorageProvider">Provider for fetching information about team details from storage table.</param>
        /// <param name="cardHelper">Instance of class that handles card create/update helper methods.</param>
        /// <param name="activityHelper">Instance of class that handles Bot activity helper methods.</param>
        /// <param name="goalHelper">Instance of class that handles goal helper methods.</param>
        /// <param name="personalGoalSearchService">Instance of class for working with personal goal data in storage.</param>
        /// <param name="userState">State management object for maintaining user conversation state.</param>
        public GoalTrackerActivityHandler(
            ILogger<GoalTrackerActivityHandler> logger,
            IStringLocalizer<Strings> localizer,
            TelemetryClient telemetryClient,
            IOptions<GoalTrackerActivityHandlerOptions> options,
            IPersonalGoalStorageProvider personalGoalStorageProvider,
            IPersonalGoalNoteStorageProvider personalGoalNoteStorageProvider,
            ITeamGoalStorageProvider teamGoalStorageProvider,
            ITeamStorageProvider teamStorageProvider,
            CardHelper cardHelper,
            ActivityHelper activityHelper,
            GoalHelper goalHelper,
            IPersonalGoalSearchService personalGoalSearchService,
            UserState userState)
        {
            this.options = options;
            this.logger = logger;
            this.localizer = localizer;
            this.telemetryClient = telemetryClient;
            this.personalGoalStorageProvider = personalGoalStorageProvider;
            this.personalGoalNoteStorageProvider = personalGoalNoteStorageProvider;
            this.teamGoalStorageProvider = teamGoalStorageProvider;
            this.teamStorageProvider = teamStorageProvider;
            this.cardHelper = cardHelper;
            this.activityHelper = activityHelper;
            this.goalHelper = goalHelper;
            this.personalGoalSearchService = personalGoalSearchService;
            this.userState = userState;
        }

        /// <summary>
        /// Invoked when a message activity is received from the user.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        /// <remarks>
        /// For more information on bot messaging in Teams, see the documentation
        /// https://docs.microsoft.com/en-us/microsoftteams/platform/bots/how-to/conversations/conversation-basics?tabs=dotnet#receive-a-message .
        /// </remarks>
        protected override async Task OnMessageActivityAsync(
            ITurnContext<IMessageActivity> turnContext,
            CancellationToken cancellationToken)
        {
            turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
            this.RecordEvent(nameof(this.OnMessageActivityAsync), turnContext);
            var activity = turnContext.Activity;

            switch (activity.Conversation.ConversationType)
            {
                case Constants.PersonalConversationType: // Command to send activities to personal bot
                    await this.OnMessageActivityInPersonalChatAsync(
                        turnContext,
                        cancellationToken);
                    break;

                case Constants.ChannelConversationType: // Command to send activities in team
                    await this.OnMessageActivityInChannelAsync(
                        activity,
                        turnContext,
                        cancellationToken);
                    break;

                default:
                    this.logger.LogInformation($"Received unexpected conversationType {activity.Conversation.ConversationType}", SeverityLevel.Warning);
                    break;
            }
        }

        /// <summary>
        /// Invoked when members other than this bot (like a user) are removed from the conversation.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        /// <remarks>
        /// Reference link: https://docs.microsoft.com/en-us/dotnet/api/microsoft.bot.builder.teams.teamsactivityhandler.onconversationupdateactivityasync?view=botbuilder-dotnet-stable
        /// </remarks>
        protected override async Task OnConversationUpdateActivityAsync(
            ITurnContext<IConversationUpdateActivity> turnContext,
            CancellationToken cancellationToken)
        {
            turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
            this.RecordEvent(nameof(this.OnConversationUpdateActivityAsync), turnContext);

            var activity = turnContext.Activity;
            var conversationType = activity.Conversation.ConversationType;
            this.logger.LogInformation("Received conversationUpdate activity");
            this.logger.LogInformation($"conversationType: {conversationType}, membersAdded: {activity.MembersAdded?.Count}, membersRemoved: {activity.MembersRemoved?.Count}");

            if (activity.MembersAdded?.Count == 0)
            {
                this.logger.LogInformation("Ignoring conversationUpdate that was not a membersAdded event");
                return;
            }

            switch (conversationType)
            {
                case Constants.PersonalConversationType: // Command to send activities to personal bot
                    await this.SendWelcomeCardInPersonalScopeAsync(turnContext, cancellationToken);
                    return;

                case Constants.ChannelConversationType: // Command to send activities in team
                    await this.OnMembersAddedOrRemovedFromTeamAsync(activity.MembersAdded, activity.MembersRemoved, turnContext, cancellationToken);
                    return;

                default:
                    this.logger.LogInformation($"Ignoring event from conversation type {conversationType}");
                    return;
            }
        }

        /// <summary>
        /// When OnTurn method receives a fetch invoke activity on bot turn, it calls this method.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="taskModuleRequest">Task module invoke request value payload.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents a task module response.</returns>
        /// <remarks>
        /// Reference link: https://docs.microsoft.com/en-us/dotnet/api/microsoft.bot.builder.teams.teamsactivityhandler.onteamstaskmodulefetchasync?view=botbuilder-dotnet-stable.
        /// </remarks>
        protected override async Task<TaskModuleResponse> OnTeamsTaskModuleFetchAsync(
            ITurnContext<IInvokeActivity> turnContext,
            TaskModuleRequest taskModuleRequest,
            CancellationToken cancellationToken)
        {
            try
            {
                turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
                this.RecordEvent(nameof(this.OnTeamsTaskModuleFetchAsync), turnContext);

                var activity = turnContext.Activity;
                var activityValue = JObject.Parse(activity.Value?.ToString())["data"].ToString();

                AdaptiveSubmitActionData adaptiveSubmitActionData = JsonConvert.DeserializeObject<AdaptiveSubmitActionData>(JObject.Parse(taskModuleRequest?.Data.ToString()).SelectToken("data").ToString());

                if (adaptiveSubmitActionData == null)
                {
                    this.logger.LogInformation("Value obtained from task module fetch action is null");
                    return this.cardHelper.GetTaskModuleErrorResponse(this.localizer.GetString("GenericErrorMessage"), this.localizer.GetString("BotFailureTitle"));
                }

                var adaptiveActionType = adaptiveSubmitActionData.AdaptiveActionType;
                switch (adaptiveActionType.ToUpperInvariant())
                {
                    case Constants.AddNoteCommand: // Command to show add note task module
                        var personalGoalDetailsForAddNote = await this.personalGoalStorageProvider.GetPersonalGoalDetailsByUserAadObjectIdAsync(activity.From.AadObjectId);
                        var personalNoteValidationResponse = this.cardHelper.GetPersonalNoteValidationResponse(personalGoalDetailsForAddNote, adaptiveSubmitActionData.GoalCycleId);
                        if (personalNoteValidationResponse == null)
                        {
                            PersonalGoalNoteDetail personalGoalNoteDetailForAddNote = JsonConvert.DeserializeObject<PersonalGoalNoteDetail>(activityValue);
                            this.logger.LogInformation("Fetch task module to show add note card");
                            return this.cardHelper.GetAddNoteCardResponse(personalGoalNoteDetailForAddNote, personalGoalDetailsForAddNote, this.localizer.GetString("AddNoteTaskModuleTitle"));
                        }

                        return personalNoteValidationResponse;

                    case Constants.EditNoteCommand: // Command to show edit note task module
                        string activityCardId = adaptiveSubmitActionData.PersonalGoalNoteId;
                        var personalGoalNoteDetail = await this.personalGoalNoteStorageProvider.GetPersonalGoalNoteDetailAsync(adaptiveSubmitActionData.PersonalGoalNoteId, activity.From.AadObjectId);
                        if (personalGoalNoteDetail == null)
                        {
                            this.logger.LogError("Value obtained for personal goal note detail from task module fetch action is null");
                            return this.cardHelper.GetTaskModuleErrorResponse(this.localizer.GetString("NoteIsDeletedErrorText"), this.localizer.GetString("EditNoteTaskModuleTitle"));
                        }

                        var personalGoalDetailsForEditNote = await this.personalGoalStorageProvider.GetPersonalGoalDetailsByUserAadObjectIdAsync(activity.From.AadObjectId);
                        var personalGoalDetails = personalGoalDetailsForEditNote.Where(personalGoal => personalGoal.PersonalGoalId == personalGoalNoteDetail.PersonalGoalId);
                        if (!personalGoalDetails.Any())
                        {
                            this.logger.LogError("Value obtained for personal goal detail is null.");
                            return this.cardHelper.GetTaskModuleErrorResponse(this.localizer.GetString("GoalIsDeletedErrorText"), this.localizer.GetString("EditNoteTaskModuleTitle"));
                        }

                        this.logger.LogInformation("Fetch task module to show card for editing note");
                        return this.cardHelper.GetAddNoteCardResponse(personalGoalNoteDetail, personalGoalDetailsForEditNote, this.localizer.GetString("EditNoteTaskModuleTitle"));

                    case Constants.SetPersonalGoalsCommand: // Command to show set personal goal card in task module
                        this.logger.LogInformation("Fetch task module to show set personal goal card");
                        return this.cardHelper.GetPersonalSetGoalCardResponse(activity.ServiceUrl, this.localizer.GetString("SetPersonalGoalTaskModuleTitle"));

                    case Constants.EditPersonalGoalsCommand: // Command to show edit personal goal card in task module
                        var goalCycleValidationResponse = await this.cardHelper.GetGoalCycleValidationResponseAsync(turnContext, adaptiveSubmitActionData.GoalCycleId, this.localizer.GetString("NoActiveGoalOrGoalCycleEndedError"), this.localizer.GetString("EditPersonalGoalTaskModuleTitle"), isPersonalGoalCycleValidation: true);
                        if (goalCycleValidationResponse == null)
                        {
                            this.logger.LogInformation("Fetch task module to show edit personal goal card");
                            return this.cardHelper.GetPersonalSetGoalCardResponse(activity.ServiceUrl, this.localizer.GetString("EditPersonalGoalTaskModuleTitle"));
                        }

                        return goalCycleValidationResponse;

                    case Constants.SetTeamGoalsCommand: // Command to show set team goal card in task module
                        this.logger.LogInformation("Fetch task module to show set team goal card");
                        return this.cardHelper.GetTeamSetGoalCardResponse(activity.ServiceUrl, this.localizer.GetString("SetTeamGoalTaskModuleTitle"));

                    case Constants.EditTeamGoalsCommand: // Command to show edit team goal card in task module
                        var goalCycleValidationResponseForTeam = await this.cardHelper.GetGoalCycleValidationResponseAsync(turnContext, adaptiveSubmitActionData.GoalCycleId, this.localizer.GetString("NoActiveGoalOrGoalCycleEndedError"), this.localizer.GetString("EditTeamGoalTaskModuleTitle"), teamId: adaptiveSubmitActionData.TeamId);
                        if (goalCycleValidationResponseForTeam == null)
                        {
                            this.logger.LogInformation("Fetch task module to show edit team goal card");
                            return this.cardHelper.GetTeamSetGoalCardResponse(activity.ServiceUrl, this.localizer.GetString("EditTeamGoalTaskModuleTitle"));
                        }

                        return goalCycleValidationResponseForTeam;

                    case Constants.AlignGoalCommand: // Command to show align goal task module
                        var teamIdForAlignGoal = adaptiveSubmitActionData.TeamId;
                        var goalCycleValidationResponseForAlignGoal = await this.cardHelper.GetGoalCycleValidationResponseAsync(turnContext, adaptiveSubmitActionData.GoalCycleId, this.localizer.GetString("NoActiveGoalOrGoalCycleEndedErrorForAlignGoal"), this.localizer.GetString("AlignGoalTaskModuleTitle"), teamId: teamIdForAlignGoal);
                        if (goalCycleValidationResponseForAlignGoal == null)
                        {
                            this.logger.LogInformation("Fetch task module to show align goal response");
                            return await this.cardHelper.GetAlignGoalTaskModuleResponseAsync(turnContext, teamIdForAlignGoal);
                        }

                        return goalCycleValidationResponseForAlignGoal;

                    default:
                        this.logger.LogInformation($"Invalid command for task module fetch activity. Command is: {adaptiveActionType}");
                        return this.cardHelper.GetTaskModuleErrorResponse(string.Format(CultureInfo.InvariantCulture, this.localizer.GetString("TaskModuleInvalidCommandText"), adaptiveActionType), this.localizer.GetString("BotFailureTitle"));
                }
            }
#pragma warning disable CA1031 // Catching general exceptions to log exception details in telemetry client and send error response in task module.
            catch (Exception ex)
#pragma warning restore CA1031 // Catching general exceptions to log exception details in telemetry client and send error response in task module.
            {
                this.logger.LogError(ex, "Error while fetching task module");
                return this.cardHelper.GetTaskModuleErrorResponse(this.localizer.GetString("GenericErrorMessage"), this.localizer.GetString("BotFailureTitle"));
            }
        }

        /// <summary>
        /// When OnTurn method receives a submit invoke activity on bot turn, it calls this method.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="taskModuleRequest">Task module invoke request value payload.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents a task module response.</returns>
        /// <remarks>
        /// Reference link: https://docs.microsoft.com/en-us/dotnet/api/microsoft.bot.builder.teams.teamsactivityhandler.onteamstaskmodulesubmitasync?view=botbuilder-dotnet-stable.
        /// </remarks>
        protected override async Task<TaskModuleResponse> OnTeamsTaskModuleSubmitAsync(
            ITurnContext<IInvokeActivity> turnContext,
            TaskModuleRequest taskModuleRequest,
            CancellationToken cancellationToken)
        {
            try
            {
                turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
                this.RecordEvent(nameof(this.OnTeamsTaskModuleSubmitAsync), turnContext);

                var activity = turnContext.Activity;
                var activityValue = JObject.Parse(activity.Value?.ToString())["data"].ToString();
                AdaptiveSubmitActionData adaptiveSubmitActionData = JsonConvert.DeserializeObject<AdaptiveSubmitActionData>(activityValue);
                if (adaptiveSubmitActionData == null)
                {
                    this.logger.LogInformation("Value obtained from task module submit action is null");
                    return this.cardHelper.GetTaskModuleErrorResponse(this.localizer.GetString("GenericErrorMessage"), this.localizer.GetString("BotFailureTitle"));
                }

                this.logger.LogInformation($"OnTeamsTaskModuleSubmitAsync: {JObject.Parse(activity.Value.ToString())["data"]}");
                var adaptiveActionType = adaptiveSubmitActionData.AdaptiveActionType;
                switch (adaptiveActionType.ToUpperInvariant())
                {
                    case Constants.AddNoteCommand: // Command to handle add/edit note task module submit action.
                        PersonalGoalNoteDetail personalGoalNoteDetail = JsonConvert.DeserializeObject<PersonalGoalNoteDetail>(activityValue);
                        var addNoteCardErrorResponse = await this.cardHelper.GetAddNoteErrorResponseAsync(personalGoalNoteDetail, turnContext);
                        if (addNoteCardErrorResponse == null)
                        {
                            await this.cardHelper.StoreOrUpdateNoteDetailAsync(turnContext, personalGoalNoteDetail, adaptiveSubmitActionData, cancellationToken);
                            this.logger.LogInformation("Value obtained from task module submit action is null");
                            return null;
                        }

                        return addNoteCardErrorResponse;

                    case Constants.SetPersonalGoalsCommand: // Command to handle set personal goal task module submit action
                        var personalGoalDetailsForSetGoal = adaptiveSubmitActionData.PersonalGoalDetails;
                        var goalCycleIdForSetGoal = adaptiveSubmitActionData.GoalCycleId;
                        var personalGoalListCardAttacment = MessageFactory.Attachment(GoalCard.GetPersonalGoalDetailsListCard(this.options.Value.AppBaseUri, personalGoalDetailsForSetGoal, this.localizer, goalCycleIdForSetGoal));
                        var setPersonalGoalCardActivityResponse = await turnContext.SendActivityAsync(personalGoalListCardAttacment, cancellationToken);
                        var isPersonalGoalsSaved = await this.goalHelper.SavePersonalGoalDetailsAsync(personalGoalDetailsForSetGoal, setPersonalGoalCardActivityResponse?.Id, turnContext);
                        if (!isPersonalGoalsSaved)
                        {
                            this.logger.LogError("Error while saving personal goal details to storage");
                            return this.cardHelper.GetTaskModuleErrorResponse(this.localizer.GetString("ErrorSavingPersonalGoalData"), this.localizer.GetString("SetPersonalGoalTaskModuleTitle"));
                        }

                        break;

                    case Constants.EditPersonalGoalsCommand: // Command to handle edit personal goal task module submit action
                        var personalGoalDetailsForEditGoal = adaptiveSubmitActionData.PersonalGoalDetails;
                        personalGoalDetailsForEditGoal = personalGoalDetailsForEditGoal.Where(personalGoalDetail => !personalGoalDetail.IsDeleted);
                        var goalCycleIdForEditGoals = adaptiveSubmitActionData.GoalCycleId;
                        var editGoalListCardAttachmentForTeam = MessageFactory.Attachment(GoalCard.GetPersonalGoalDetailsListCard(this.options.Value.AppBaseUri, personalGoalDetailsForEditGoal, this.localizer, goalCycleIdForEditGoals));

                        // Update goals list card in personal chat.
                        editGoalListCardAttachmentForTeam.Id = adaptiveSubmitActionData.ActivityCardId;
                        this.logger.LogInformation("Updating personal goal list card");
                        var editGoalCardActivityResponse = await turnContext.UpdateActivityAsync(editGoalListCardAttachmentForTeam, cancellationToken);

                        // Update storage with personal goal details.
                        var isPersonalGoalUpdated = await this.goalHelper.SavePersonalGoalDetailsAsync(personalGoalDetailsForEditGoal, editGoalCardActivityResponse?.Id, turnContext);
                        if (!isPersonalGoalUpdated)
                        {
                            this.logger.LogError("Error while saving personal goal details to storage");
                            return this.cardHelper.GetTaskModuleErrorResponse(this.localizer.GetString("ErrorSavingPersonalGoalData"), this.localizer.GetString("SetPersonalGoalTaskModuleTitle"));
                        }

                        break;

                    case Constants.SetTeamGoalsCommand: // Command to handle set team goal task module submit action
                        var teamGoalDetailsForSetGoal = adaptiveSubmitActionData.TeamGoalDetails;
                        var goalCycleIdForSetTeamGoal = adaptiveSubmitActionData.GoalCycleId;
                        var teamGoalListCardAttachmentForTeam = MessageFactory.Attachment(GoalCard.GetTeamGoalDetailsListCard(this.options.Value.AppBaseUri, adaptiveSubmitActionData.TeamGoalDetails, this.localizer, activity.From.Name, goalCycleIdForSetTeamGoal));

                        // Sending team goal list card as a separate message in team
                        turnContext.Activity.Conversation.Id = turnContext.Activity.Conversation.Id.Split(';')[0];
                        this.logger.LogInformation("Sending team goal card in team.");
                        var setTeamGoalCardActivityResponse = await turnContext.SendActivityAsync(teamGoalListCardAttachmentForTeam, cancellationToken);

                        // Update storage with team goal details.
                        this.logger.LogInformation("Saving team goal details data in storage");
                        var isTeamGoalSaved = await this.goalHelper.SaveTeamGoalDetailsAsync(teamGoalDetailsForSetGoal, setTeamGoalCardActivityResponse?.Id, turnContext);
                        if (!isTeamGoalSaved)
                        {
                            this.logger.LogError("Error while saving personal goal details to storage");
                            return this.cardHelper.GetTaskModuleErrorResponse(this.localizer.GetString("ErrorSavingPersonalGoalData"), this.localizer.GetString("SetPersonalGoalTaskModuleTitle"));
                        }

                        // Send List card with team goal details to all members of the team.
                        await this.cardHelper.SendTeamGoalListCardToTeamMembersAsync(this.options.Value.AppBaseUri, turnContext, teamGoalDetailsForSetGoal, goalCycleIdForSetTeamGoal, cancellationToken);
                        break;

                    case Constants.EditTeamGoalsCommand: // Command to handle edit team goal task module submit action
                        var goalCycleIdForEditTeamGoals = adaptiveSubmitActionData.GoalCycleId;
                        var teamGoalDetailsForEditGoal = adaptiveSubmitActionData.TeamGoalDetails;
                        teamGoalDetailsForEditGoal = teamGoalDetailsForEditGoal.Where(teamGoal => !teamGoal.IsDeleted);
                        var editTeamGoalListCardAttachmentForTeam = MessageFactory.Attachment(GoalCard.GetTeamGoalDetailsListCard(this.options.Value.AppBaseUri, teamGoalDetailsForEditGoal, this.localizer, activity.From.Name, goalCycleIdForEditTeamGoals));

                        // Update goals list card in personal chat.
                        editTeamGoalListCardAttachmentForTeam.Id = adaptiveSubmitActionData.ActivityCardId;
                        this.logger.LogInformation("Update existing team goal list card of team");
                        var editTeamGoalCardActivityResponse = await turnContext.UpdateActivityAsync(editTeamGoalListCardAttachmentForTeam, cancellationToken);

                        this.logger.LogInformation("Updating team goal details data in storage");
                        var isTeamGoalUpdated = await this.goalHelper.SaveTeamGoalDetailsAsync(teamGoalDetailsForEditGoal, editTeamGoalCardActivityResponse?.Id, turnContext);
                        if (!isTeamGoalUpdated)
                        {
                            this.logger.LogError("Error while saving team goal details to storage");
                            return this.cardHelper.GetTaskModuleErrorResponse(this.localizer.GetString("ErrorSavingTeamGoalData"), this.localizer.GetString("SetTeamGoalTaskModuleTitle"));
                        }

                        // Send List card with team goal details to all members of the team.
                        await this.cardHelper.SendTeamGoalListCardToTeamMembersAsync(this.options.Value.AppBaseUri, turnContext, teamGoalDetailsForEditGoal, goalCycleIdForEditTeamGoals, cancellationToken);
                        break;

                    case Constants.AlignGoalCommand: // Command to handle align goal task module submit action
                        this.logger.LogInformation("Fetch task module to show align goal response");
                        return await this.cardHelper.GetAlignGoalTaskModuleResponseAsync(turnContext.Activity.From.AadObjectId, adaptiveSubmitActionData.TeamId);

                    default:
                        this.logger.LogInformation($"Invalid command for task module submit activity. Command is: {adaptiveActionType}");
                        return this.cardHelper.GetTaskModuleErrorResponse(string.Format(CultureInfo.InvariantCulture, this.localizer.GetString("TaskModuleInvalidCommandText"), adaptiveActionType), this.localizer.GetString("BotFailureTitle"));
                }

                return null;
            }
#pragma warning disable CA1031 // Catching general exceptions to log exception details in telemetry client and send error response in task module.
            catch (Exception ex)
#pragma warning restore CA1031 // Catching general exceptions to log exception details in telemetry client and send error response in task module.
            {
                this.logger.LogError(ex, "Error while submitting task module");
                return this.cardHelper.GetTaskModuleErrorResponse(this.localizer.GetString("GenericErrorMessage"), this.localizer.GetString("BotFailureTitle"));
            }
        }

        /// <summary>
        /// Handle message extension action fetch task received by the bot.
        /// </summary>
        /// <param name="turnContext">Provides context for a turn of a bot.</param>
        /// <param name="action">Messaging extension action value payload.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>Response of messaging extension action.</returns>
        /// <remarks>
        /// Reference link: https://docs.microsoft.com/en-us/dotnet/api/microsoft.bot.builder.teams.teamsactivityhandler.onteamsmessagingextensionfetchtaskasync?view=botbuilder-dotnet-stable.
        /// </remarks>
        protected override async Task<MessagingExtensionActionResponse> OnTeamsMessagingExtensionFetchTaskAsync(
            ITurnContext<IInvokeActivity> turnContext,
            MessagingExtensionAction action,
            CancellationToken cancellationToken)
        {
            try
            {
                turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
                action = action ?? throw new ArgumentNullException(nameof(action));
                this.RecordEvent(nameof(this.OnTeamsMessagingExtensionFetchTaskAsync), turnContext);

                var actionType = action.CommandId;
                if (actionType == null)
                {
                    this.logger.LogInformation("Value obtained from task module fetch message action is null");
                    return await this.cardHelper.GetTaskModuleErrorResponseAsync(this.localizer.GetString("GenericErrorMessage"), this.localizer.GetString("BotFailureTitle"));
                }

                switch (actionType.ToUpperInvariant())
                {
                    case Constants.AddNoteCommand: // Command text to fetch add note task module through message action
                        if (action.MessagePayload.Body.ContentType == ContentTypeHtml)
                        {
                            this.logger.LogInformation($"Add note command cannot be invoked for bot messages");
                            return await this.cardHelper.GetTaskModuleErrorResponseAsync(this.localizer.GetString("AddNoteNotAllowedForBotMessage"), this.localizer.GetString("AddNoteTaskModuleTitle"));
                        }

                        PersonalGoalNoteDetail personalGoalNoteDetail = new PersonalGoalNoteDetail
                        {
                            PersonalGoalNoteDescription = action.MessagePayload.Body.Content,
                            SourceName = action.MessagePayload.From.User.DisplayName,
                        };
                        var personalGoalDetail = await this.personalGoalStorageProvider.GetPersonalGoalDetailsByUserAadObjectIdAsync(turnContext.Activity.From.AadObjectId);
                        this.logger.LogInformation("Fetch task module to show add note card invoked through message action");
                        return await this.cardHelper.GetAddNoteCardResponseAsync(personalGoalNoteDetail, personalGoalDetail);

                    default:
                        this.logger.LogInformation($"Invalid command for task module fetch activity. Command is: {actionType}");
                        return await this.cardHelper.GetTaskModuleErrorResponseAsync(string.Format(CultureInfo.InvariantCulture, this.localizer.GetString("TaskModuleInvalidCommandText"), actionType), this.localizer.GetString("BotFailureTitle"));
                }
            }
#pragma warning disable CA1031 // Catching general exceptions to log exception details in telemetry client and send error response in task module.
            catch (Exception ex)
#pragma warning restore CA1031 // Catching general exceptions to log exception details in telemetry client and send error response in task module.
            {
                this.logger.LogError(ex, "Error while fetching task module through message action");
                return await this.cardHelper.GetTaskModuleErrorResponseAsync(this.localizer.GetString("GenericErrorMessage"), this.localizer.GetString("AddNoteTaskModuleTitle"));
            }
        }

        /// <summary>
        ///  Handle message extension submit action task received by the bot.
        /// </summary>
        /// <param name="turnContext">Provides context for a turn of a bot.</param>
        /// <param name="action">Messaging extension action value payload.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>Response of messaging extension action.</returns>
        /// <remarks>
        /// Reference link: https://docs.microsoft.com/en-us/dotnet/api/microsoft.bot.builder.teams.teamsactivityhandler.onteamsmessagingextensionfetchtaskasync?view=botbuilder-dotnet-stable.
        /// </remarks>
        protected override async Task<MessagingExtensionActionResponse> OnTeamsMessagingExtensionSubmitActionAsync(
            ITurnContext<IInvokeActivity> turnContext,
            MessagingExtensionAction action,
            CancellationToken cancellationToken)
        {
            try
            {
                turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
                action = action ?? throw new ArgumentNullException(nameof(action));
                this.RecordEvent(nameof(this.OnTeamsMessagingExtensionFetchTaskAsync), turnContext);

                var activity = turnContext.Activity;
                var activityValue = JObject.Parse(activity.Value?.ToString())["data"].ToString();
                AdaptiveSubmitActionData adaptiveSubmitActionData = JsonConvert.DeserializeObject<AdaptiveSubmitActionData>(activityValue);
                if (adaptiveSubmitActionData == null)
                {
                    this.logger.LogInformation("Value obtained from task module submit using message action is null");
                    return await this.cardHelper.GetTaskModuleErrorResponseAsync(this.localizer.GetString("GenericErrorMessage"), this.localizer.GetString("BotFailureTitle"));
                }

                var actionType = adaptiveSubmitActionData.AdaptiveActionType;
                switch (actionType.ToUpperInvariant())
                {
                    case Constants.AddNoteCommand: // Command to handle add note task module submit action for message action.
                        PersonalGoalNoteDetail personalGoalNoteDetail = JsonConvert.DeserializeObject<PersonalGoalNoteDetail>(activityValue);
                        var addNoteErrorResponse = await this.cardHelper.GetAddNoteErrorResponseForMessageActionAsync(personalGoalNoteDetail, turnContext);
                        if (addNoteErrorResponse == null)
                        {
                            await this.cardHelper.StoreOrUpdateNoteDetailAsync(turnContext, personalGoalNoteDetail, adaptiveSubmitActionData, cancellationToken);
                            this.logger.LogInformation($"{nameof(this.OnTeamsMessagingExtensionSubmitActionAsync)}: {JObject.Parse(activity.Value.ToString())["data"]}");
                            return null;
                        }

                        return addNoteErrorResponse;

                    default:
                        this.logger.LogInformation($"Invalid command for task module submit activity. Command is: {actionType}");
                        return await this.cardHelper.GetTaskModuleErrorResponseAsync(string.Format(CultureInfo.InvariantCulture, this.localizer.GetString("TaskModuleInvalidCommandText"), actionType), this.localizer.GetString("BotFailureTitle"));
                }
            }
#pragma warning disable CA1031 // Catching general exceptions to log exception details in telemetry client.
            catch (Exception ex)
#pragma warning restore CA1031 // Catching general exceptions to log exception details in telemetry client.
            {
                this.logger.LogError(ex, "Error while submitting task module received by the bot which is invoked through message action");
                return await this.cardHelper.GetTaskModuleErrorResponseAsync(this.localizer.GetString("GenericErrorMessage"), this.localizer.GetString("AddNoteTaskModuleTitle"));
            }
        }

        /// <summary>
        /// Handle message activity in 1:1 chat.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task OnMessageActivityInPersonalChatAsync(
        ITurnContext<IMessageActivity> turnContext,
        CancellationToken cancellationToken)
        {
            string text = (turnContext.Activity.Text ?? string.Empty).Trim().ToUpperInvariant();
            switch (text)
            {
                case Constants.SetGoalsCommand: // Command to send card with set goal button
                    var setPersonalGoalAttachment = MessageFactory.Attachment(GoalCard.GetSetPersonalGoalsCard(this.localizer));
                    this.logger.LogInformation("Sending card for setting personal goal");
                    await turnContext.SendActivityAsync(setPersonalGoalAttachment, cancellationToken);
                    break;

                case Constants.EditGoalsCommand: // Command to send card with edit goal button
                    var editPersonalGoalAttachment = MessageFactory.Attachment(GoalCard.GetEditPersonalGoalsCard(this.localizer));
                    this.logger.LogInformation("Sending card for editing personal goal");
                    await turnContext.SendActivityAsync(editPersonalGoalAttachment, cancellationToken);
                    break;

                case Constants.AddNoteCommand: // Command to send card with add note button
                    var addNoteAttachment = MessageFactory.Attachment(NoteCard.GetAddNoteCardOnMessage(this.localizer));
                    this.logger.LogInformation("Sending card for adding note in personal goal");
                    await turnContext.SendActivityAsync(addNoteAttachment, cancellationToken);
                    break;

                default:
                    this.logger.LogInformation("Invalid command text entered in personal chat. Sending help card");
                    var helpCardAttachment = MessageFactory.Attachment(HelpCard.GetHelpCardInPersonalChat(this.localizer));
                    await turnContext.SendActivityAsync(helpCardAttachment, cancellationToken);
                    break;
            }
        }

        /// <summary>
        /// Handle message activity in channel.
        /// </summary>
        /// <param name="message">A message in a conversation.</param>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task OnMessageActivityInChannelAsync(
            IMessageActivity message,
            ITurnContext<IMessageActivity> turnContext,
            CancellationToken cancellationToken)
        {
            string actionType = message.Value != null ? JObject.Parse(message.Value.ToString())["AdaptiveActionType"]?.ToString() : null;
            message.RemoveRecipientMention();
            string text = string.IsNullOrEmpty(message.Text) ? actionType : message.Text.Trim();
            switch (text.ToUpperInvariant())
            {
                case Constants.SetTeamGoalsCommand: // Command to send card with set goal button
                    this.logger.LogInformation("Sending card for setting team goal");
                    var setTeamGoalCardAttachment = MessageFactory.Attachment(GoalCard.GetSetTeamGoalsCard(this.localizer));
                    await turnContext.SendActivityAsync(setTeamGoalCardAttachment, cancellationToken);
                    break;

                case Constants.EditTeamGoalsCommand: // Command to send card with edit goal button
                    this.logger.LogInformation("Sending card for editing team goal");
                    var editTeamGoalCardAttachment = MessageFactory.Attachment(GoalCard.GetEditTeamGoalsCard(this.localizer, turnContext.Activity.TeamsGetTeamInfo().Id));
                    await turnContext.SendActivityAsync(editTeamGoalCardAttachment, cancellationToken);
                    break;

                case Constants.GoalStatusCommand: // Command to send goal status card
                    var goalStatusCardAttachment = await this.activityHelper.GetGoalStatusAttachmentAsync(turnContext, turnContext.Activity.TeamsGetTeamInfo().Id, this.options.Value.AppBaseUri, cancellationToken);
                    if (goalStatusCardAttachment != null)
                    {
                        this.logger.LogInformation("Sending goal status card in team");
                        await turnContext.SendActivityAsync(goalStatusCardAttachment, cancellationToken);
                    }

                    break;

                default:
                    this.logger.LogInformation("Invalid command text entered in channel. Sending help card");
                    var helpCardAttachment = MessageFactory.Attachment(HelpCard.GetHelpCardInChannel(this.localizer));
                    await turnContext.SendActivityAsync(helpCardAttachment, cancellationToken);
                    break;
            }
        }

        /// <summary>
        /// Handle 1:1 chat with members who started chat for the first time.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task SendWelcomeCardInPersonalScopeAsync(
            ITurnContext<IConversationUpdateActivity> turnContext,
            CancellationToken cancellationToken)
        {
            this.logger.LogInformation($"Bot added in personal {turnContext.Activity.Conversation.Id}");
            var userStateAccessors = this.userState.CreateProperty<UserConversationState>(nameof(UserConversationState));
            var userConversationState = await userStateAccessors.GetAsync(turnContext, () => new UserConversationState());

            if (userConversationState == null || !userConversationState.IsWelcomeCardSent)
            {
                userConversationState.IsWelcomeCardSent = true;
                await userStateAccessors.SetAsync(turnContext, userConversationState);
                var welcomeCardAttachment = MessageFactory.Attachment(WelcomeCard.GetWelcomeCardAttachmentForPersonalChat(this.options.Value.AppBaseUri, this.localizer));
                this.logger.LogInformation($"Sending welcome card to user in personal chat.");
                await turnContext.SendActivityAsync(welcomeCardAttachment, cancellationToken);
            }
        }

        /// <summary>
        /// Handle members added conversationUpdate event in team.
        /// </summary>
        /// <param name="membersAdded">Channel account information needed to route a message.</param>
        /// <param name="membersRemoved">Channel account information needed to route a message</param>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task OnMembersAddedOrRemovedFromTeamAsync(
           IList<ChannelAccount> membersAdded,
           IList<ChannelAccount> membersRemoved,
           ITurnContext<IConversationUpdateActivity> turnContext,
           CancellationToken cancellationToken)
        {
            var activity = turnContext.Activity;
            var teamsDetails = activity.TeamsGetTeamInfo();

            if (membersAdded != null && membersAdded.Any(channelAccount => channelAccount.Id == activity.Recipient.Id))
            {
                // Bot was added to a team.
                this.logger.LogInformation($"Bot added to team {activity.Conversation.Id}. Sending welcome card in team.");

                // Storing team information to storage
                TeamDetail teamEntity = new TeamDetail
                {
                    TeamId = teamsDetails.Id,
                    BotInstalledOn = DateTime.UtcNow,
                    ServiceUrl = turnContext.Activity.ServiceUrl,
                };

                bool operationStatus = await this.teamStorageProvider.StoreOrUpdateTeamDetailAsync(teamEntity);
                if (!operationStatus)
                {
                    this.logger.LogInformation($"Unable to store bot Installation detail in table storage.");
                }

                var welcomeCardAttachment = MessageFactory.Attachment(WelcomeCard.GetWelcomeCardAttachmentForChannel(this.localizer));
                await turnContext.SendActivityAsync(welcomeCardAttachment, cancellationToken);
            }

            if (membersRemoved != null && membersRemoved.Any(channelAccount => channelAccount.Id == activity.Recipient.Id))
            {
                var teamEntity = await this.teamStorageProvider.GetTeamDetailAsync(teamsDetails.Id);
                bool operationStatus = await this.teamStorageProvider.DeleteTeamDetailAsync(teamEntity);
                if (!operationStatus)
                {
                    this.logger.LogInformation($"Unable to remove team details from table storage.");
                }

                foreach (var member in membersRemoved)
                {
                    var teamId = turnContext.Activity.Conversation.Id;
                    var personalGoalDetails = await this.personalGoalStorageProvider.GetUserAlignedGoalDetailsByTeamIdAsync(turnContext.Activity.TeamsGetTeamInfo().Id, member.AadObjectId);
                    if (personalGoalDetails != null && personalGoalDetails.Any())
                    {
                        // Unaligned team goals from personal goal details when user leaves the team or gets removed from the team:
                        await this.goalHelper.DeleteAlignedGoalDetailsAsync(personalGoalDetails);
                    }
                }
            }
        }

        /// <summary>
        /// Records event occurred in the application in Application Insights telemetry client.
        /// </summary>
        /// <param name="eventName"> Name of the event.</param>
        /// <param name="turnContext"> Context object containing information cached for a single turn of conversation with a user.</param>
        private void RecordEvent(string eventName, ITurnContext turnContext)
        {
            var teamsChannelData = turnContext.Activity.GetChannelData<TeamsChannelData>();

            this.telemetryClient.TrackEvent(eventName, new Dictionary<string, string>
            {
                { "userId", turnContext.Activity.From.AadObjectId },
                { "tenantId", turnContext.Activity.Conversation.TenantId },
                { "teamId", teamsChannelData?.Team?.Id },
                { "channelId", teamsChannelData?.Channel?.Id },
            });
        }
    }
}