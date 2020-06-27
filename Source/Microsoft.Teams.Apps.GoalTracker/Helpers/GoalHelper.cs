// <copyright file="GoalHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Teams;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Localization;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.GoalTracker.Common;
    using Microsoft.Teams.Apps.GoalTracker.Models;

    /// <summary>
    /// Class that handles goal helper methods.
    /// </summary>
    public class GoalHelper
    {
        /// <summary>
        /// Instance to send logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<GoalHelper> logger;

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
        /// Initializes a new instance of the <see cref="GoalHelper"/> class.
        /// </summary>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="options">A set of key/value application configuration properties for activity handler.</param>
        /// <param name="personalGoalStorageProvider">Storage provider for working with personal goal data in storage.</param>
        /// <param name="personalGoalNoteStorageProvider">Storage provider for working with personal goal note data in storage</param>
        /// <param name="teamGoalStorageProvider">Storage provider for working with team goal data in storage.</param>
        public GoalHelper(
            ILogger<GoalHelper> logger,
            IStringLocalizer<Strings> localizer,
            IOptions<GoalTrackerActivityHandlerOptions> options,
            IPersonalGoalStorageProvider personalGoalStorageProvider,
            IPersonalGoalNoteStorageProvider personalGoalNoteStorageProvider,
            ITeamGoalStorageProvider teamGoalStorageProvider)
        {
            this.logger = logger;
            this.localizer = localizer;
            this.options = options;
            this.personalGoalStorageProvider = personalGoalStorageProvider;
            this.personalGoalNoteStorageProvider = personalGoalNoteStorageProvider;
            this.teamGoalStorageProvider = teamGoalStorageProvider;
        }

        /// <summary>
        /// Delete aligned goal details data from storage when user wants to change goal alignment from one team to another.
        /// </summary>
        /// <param name="personalGoalDetails">Holds collection of personal goal note detail entity data.</param>
        /// <returns>Return true for successful operation.</returns>
        public async Task<bool> DeleteAlignedGoalDetailsAsync(IEnumerable<PersonalGoalDetail> personalGoalDetails)
        {
            personalGoalDetails = personalGoalDetails ?? throw new ArgumentNullException(nameof(personalGoalDetails));

            try
            {
                foreach (var personalGoalDetail in personalGoalDetails)
                {
                    personalGoalDetail.IsAligned = false;
                    personalGoalDetail.LastModifiedOn = DateTime.UtcNow.ToString(CultureInfo.InvariantCulture);
                    personalGoalDetail.TeamId = null;
                    personalGoalDetail.TeamGoalId = null;
                }

                return await this.personalGoalStorageProvider.CreateOrUpdatePersonalGoalDetailsAsync(personalGoalDetails);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"An error occurred in {nameof(this.DeleteAlignedGoalDetailsAsync)} while deleting aligned goals.");
                throw;
            }
        }

        /// <summary>
        /// Method to store team goal details in storage
        /// </summary>
        /// <param name="teamGoalDetails">Team goal detail entities received from client application.</param>
        /// <param name="setTeamGoalCardActivityResponseId">Response id of list card sent in channel.</param>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <returns>A task that stores team goal details to storage.</returns>
        public async Task<bool> SaveTeamGoalDetailsAsync(IEnumerable<TeamGoalDetail> teamGoalDetails, string setTeamGoalCardActivityResponseId, ITurnContext<IInvokeActivity> turnContext)
        {
            try
            {
                turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
                teamGoalDetails = teamGoalDetails ?? throw new ArgumentNullException(nameof(teamGoalDetails));
                var activity = turnContext.Activity;

                // Store team goal details to storage
                foreach (var teamGoal in teamGoalDetails)
                {
                    teamGoal.LastModifiedOn = DateTime.UtcNow.ToString(CultureInfo.CurrentCulture);
                    teamGoal.LastModifiedBy = activity.From.Name;
                    teamGoal.AdaptiveCardActivityId = setTeamGoalCardActivityResponseId;
                    teamGoal.ChannelConversationId = activity.Conversation.Id;
                }

                return await this.teamGoalStorageProvider.CreateOrUpdateTeamGoalDetailsAsync(teamGoalDetails);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"An error occurred in {nameof(this.SaveTeamGoalDetailsAsync)} while saving team goal details with list card id: {setTeamGoalCardActivityResponseId}");
                throw;
            }
        }

        /// <summary>
        /// Method to store personal goal details in storage
        /// </summary>
        /// <param name="personalGoalDetails">Personal goal detail entities received from client application.</param>
        /// <param name="setPersonalGoalCardActivityResponseId">Response id of personal goal list card.</param>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <returns>A task that stores personal goal details to storage. </returns>
        public async Task<bool> SavePersonalGoalDetailsAsync(IEnumerable<PersonalGoalDetail> personalGoalDetails, string setPersonalGoalCardActivityResponseId, ITurnContext<IInvokeActivity> turnContext)
        {
            try
            {
                turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
                personalGoalDetails = personalGoalDetails ?? throw new ArgumentNullException(nameof(personalGoalDetails));

                var activity = turnContext.Activity;

                // Store personal goal details to storage
                foreach (var personalGoalDetail in personalGoalDetails)
                {
                    personalGoalDetail.LastModifiedOn = DateTime.UtcNow.ToString(CultureInfo.CurrentCulture);
                    personalGoalDetail.LastModifiedBy = activity.From.Name;
                    personalGoalDetail.AdaptiveCardActivityId = setPersonalGoalCardActivityResponseId;
                    personalGoalDetail.ConversationId = activity.Conversation.Id;
                }

                return await this.personalGoalStorageProvider.CreateOrUpdatePersonalGoalDetailsAsync(personalGoalDetails);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"An error occurred in {nameof(this.SavePersonalGoalDetailsAsync)} while saving personal goal details with list card id: {setPersonalGoalCardActivityResponseId}");
                throw;
            }
        }

        /// <summary>
        /// Get list of members present in the team.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>Returns list of members present in the team.</returns>
        public async Task<IEnumerable<TeamsChannelAccount>> GetMembersInTeamAsync(ITurnContext turnContext, CancellationToken cancellationToken)
        {
            try
            {
                var teamsChannelAccounts = new List<TeamsChannelAccount>();
                string continuationToken = null;
                do
                {
                    var currentPage = await TeamsInfo.GetPagedMembersAsync(turnContext, 100, continuationToken, cancellationToken);
                    continuationToken = currentPage.ContinuationToken;
                    teamsChannelAccounts.AddRange(currentPage.Members);
                }
                while (continuationToken != null);

                return teamsChannelAccounts;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while obtaining team members of the team.");
                throw;
            }
        }
    }
}
