// <copyright file="ActivityHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Schema;
    using Microsoft.Extensions.Localization;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.GoalTracker.Cards;
    using Microsoft.Teams.Apps.GoalTracker.Common;
    using Microsoft.Teams.Apps.GoalTracker.Models;

    /// <summary>
    /// Instance of class that handles Bot activity helper methods.
    /// </summary>
    public class ActivityHelper
    {
        /// <summary>
        /// Instance to send logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<ActivityHelper> logger;

        /// <summary>
        /// The current cultures' string localizer.
        /// </summary>
        private readonly IStringLocalizer<Strings> localizer;

        /// <summary>
        /// Storage provider for working with team goal data in storage.
        /// </summary>
        private readonly ITeamGoalStorageProvider teamGoalStorageProvider;

        /// <summary>
        /// Instance of search service for working with personal goal data in storage.
        /// </summary>
        private readonly IPersonalGoalSearchService personalGoalSearchService;

        /// <summary>
        /// Initializes a new instance of the <see cref="ActivityHelper"/> class.
        /// </summary>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="teamGoalStorageProvider">Storage provider for working with team goal data in storage.</param>
        /// <param name="personalGoalSearchService">Personal goal search service which will help in retrieving aligned goals information.</param>
        public ActivityHelper(
            ILogger<ActivityHelper> logger,
            IStringLocalizer<Strings> localizer,
            ITeamGoalStorageProvider teamGoalStorageProvider,
            IPersonalGoalSearchService personalGoalSearchService)
        {
            this.logger = logger;
            this.localizer = localizer;
            this.teamGoalStorageProvider = teamGoalStorageProvider;
            this.personalGoalSearchService = personalGoalSearchService;
        }

        /// <summary>
        /// Get goal status attachment.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="teamId">Team id to get teams goal details from storage.</param>
        /// <param name="applicationBasePath">Application base path.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>Return goal status card as attachment.</returns>
        public async Task<IMessageActivity> GetGoalStatusAttachmentAsync(ITurnContext<IMessageActivity> turnContext, string teamId, string applicationBasePath, CancellationToken cancellationToken)
        {
            try
            {
                turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));

                var teamGoalDetails = await this.teamGoalStorageProvider.GetTeamGoalDetailsByTeamIdAsync(teamId);
                if (teamGoalDetails == null)
                {
                    await turnContext.SendActivityAsync(this.localizer.GetString("GoalStatusNoTeamGoalAvailable"), cancellationToken: cancellationToken);
                    return null;
                }

                var notStartedGoals = await this.personalGoalSearchService.SearchPersonalGoalsWithStatusAsync(PersonalGoalSearchScope.NotStarted, string.Empty, teamId);
                var inProgressGoals = await this.personalGoalSearchService.SearchPersonalGoalsWithStatusAsync(PersonalGoalSearchScope.InProgress, string.Empty, teamId);
                var completedGoals = await this.personalGoalSearchService.SearchPersonalGoalsWithStatusAsync(PersonalGoalSearchScope.Completed, string.Empty, teamId);

                var teamGoalStatuses = this.MergeTeamGoalStatusCollections(teamGoalDetails, notStartedGoals, inProgressGoals, completedGoals);

                teamGoalStatuses = teamGoalStatuses.Where(teamGoalStatus => teamGoalStatus.NotStartedGoalCount != 0
                || teamGoalStatus.InProgressGoalCount != 0 || teamGoalStatus.CompletedGoalCount != 0).ToList();

                if (teamGoalStatuses == null || !teamGoalStatuses.Any())
                {
                    await turnContext.SendActivityAsync(this.localizer.GetString("GoalStatusNoGoalsAligned"), cancellationToken: cancellationToken);
                    return null;
                }

                turnContext.Activity.Conversation.Id = turnContext.Activity.Conversation.Id.Split(";")[0];
                return MessageFactory.Attachment(GoalStatusCard.GetGoalStatusCard(teamGoalStatuses, teamGoalDetails.First().TeamGoalStartDate, teamGoalDetails.First().TeamGoalEndDate, applicationBasePath, this.localizer));
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"An error occurred in {nameof(this.GetGoalStatusAttachmentAsync)} while fetching team goal status for : {teamId}");
                throw;
            }
        }

        /// <summary>
        /// Handles status counts of personal goals which are aligned to multiple team goals.
        /// Method handles comma separated team goal ids from collection and creates individual object for each team goal and add in collection.
        /// </summary>
        /// <param name="teamGoalStatuses">Collection of team goal status containing not started/in progress/completed status counts for each aligned team goal</param>
        /// <returns>Returns collection of team goal status containing goal status data about each aligned team goal.</returns>
        private IEnumerable<TeamGoalStatus> HandleGoalsWithMultipleTeamGoalsAligned(IEnumerable<TeamGoalStatus> teamGoalStatuses)
        {
            if (!teamGoalStatuses.Any(teamGoalStatus => teamGoalStatus.TeamGoalId.Contains(',', StringComparison.OrdinalIgnoreCase)))
            {
                return teamGoalStatuses;
            }

            var teamGoalStatusDetails = teamGoalStatuses.Where(teamGoalStatus => !teamGoalStatus.TeamGoalId.Contains(',', StringComparison.OrdinalIgnoreCase)).ToList();
            var multipleTeamGoalStatuses = teamGoalStatuses.Where(teamGoalStatus => teamGoalStatus.TeamGoalId.Contains(',', StringComparison.OrdinalIgnoreCase)).ToList();

            foreach (var teamGoalStatus in multipleTeamGoalStatuses)
            {
                var teamgoalIds = teamGoalStatus.TeamGoalId.Split(',');
                foreach (var teamGoalId in teamgoalIds)
                {
                    var teamGoalStatusDetail = teamGoalStatusDetails.Where(teamGoalStatus => teamGoalStatus.TeamGoalId == teamGoalId.Trim()).FirstOrDefault();
                    if (teamGoalStatusDetail != null)
                    {
                        // If team goal id already exists in collection, then add status counts to existing team goal id.
                        teamGoalStatusDetails.ForEach(teamGoalStatusDetail =>
                        {
                            if (teamGoalStatusDetail.TeamGoalId == teamGoalId)
                            {
                                teamGoalStatusDetail.NotStartedGoalCount += teamGoalStatus.NotStartedGoalCount;
                                teamGoalStatusDetail.InProgressGoalCount += teamGoalStatus.InProgressGoalCount;
                                teamGoalStatusDetail.CompletedGoalCount += teamGoalStatus.CompletedGoalCount;
                            }
                        });
                    }
                    else
                    {
                        // Create new object for current team goal id, add status counts and add object in collection.
                        TeamGoalStatus goalStatusDetail = new TeamGoalStatus
                        {
                            NotStartedGoalCount = teamGoalStatus.NotStartedGoalCount,
                            InProgressGoalCount = teamGoalStatus.InProgressGoalCount,
                            CompletedGoalCount = teamGoalStatus.CompletedGoalCount,
                            TeamGoalId = teamGoalId.Trim(),
                        };
                        teamGoalStatusDetails.Add(goalStatusDetail);
                    }
                }
            }

            return teamGoalStatusDetails;
        }

        /// <summary>
        /// Merge goal status collection of team goals which are aligned to personal goals.
        /// </summary>
        /// <param name="teamGoalDetails">Collection of team goal details added for particular team.</param>
        /// <param name="notStartedGoals">Collection of team goal status containing not started status counts for each aligned team goal.</param>
        /// <param name="inProgressGoals">Collection of team goal status containing in progress status counts for each aligned team goal.</param>
        /// <param name="completedGoals">Collection of team goal status containing completed status counts for each aligned team goal.</param>
        /// <returns>Returns collection of team goal status containing goal status data about each aligned team goal.</returns>
        private IEnumerable<TeamGoalStatus> MergeTeamGoalStatusCollections(
            IEnumerable<TeamGoalDetail> teamGoalDetails,
            IEnumerable<TeamGoalStatus> notStartedGoals,
            IEnumerable<TeamGoalStatus> inProgressGoals,
            IEnumerable<TeamGoalStatus> completedGoals)
        {
            notStartedGoals = this.HandleGoalsWithMultipleTeamGoalsAligned(notStartedGoals);
            inProgressGoals = this.HandleGoalsWithMultipleTeamGoalsAligned(inProgressGoals);
            completedGoals = this.HandleGoalsWithMultipleTeamGoalsAligned(completedGoals);

            // Merge not started/in progress/completed goal status lists into single list.
            var teamGoalStatuses = (from teamGoalCollection in teamGoalDetails
                                    join notStartedGoalCollection in notStartedGoals on teamGoalCollection.TeamGoalId equals notStartedGoalCollection.TeamGoalId into notStartedGoalDetails
                                    join inProgressGoalCollection in inProgressGoals on teamGoalCollection.TeamGoalId equals inProgressGoalCollection.TeamGoalId into inProgressGoalDetails
                                    join completedGoalCollection in completedGoals on teamGoalCollection.TeamGoalId equals completedGoalCollection.TeamGoalId into completedGoalDetails
                                    select new TeamGoalStatus
                                    {
                                        TeamGoalId = teamGoalCollection.TeamGoalId,
                                        TeamGoalName = teamGoalCollection.TeamGoalName,
                                        NotStartedGoalCount = notStartedGoalDetails.Any() ? notStartedGoalDetails.First().NotStartedGoalCount : 0,
                                        InProgressGoalCount = inProgressGoalDetails.Any() ? inProgressGoalDetails.First().InProgressGoalCount : 0,
                                        CompletedGoalCount = completedGoalDetails.Any() ? completedGoalDetails.First().CompletedGoalCount : 0,
                                    }).ToList();

            return teamGoalStatuses;
        }
    }
}