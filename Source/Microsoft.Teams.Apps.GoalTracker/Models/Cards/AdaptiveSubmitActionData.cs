// <copyright file="AdaptiveSubmitActionData.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Models
{
    using System.Collections.Generic;
    using Microsoft.Bot.Schema;
    using Newtonsoft.Json;

    /// <summary>
    /// Adaptive submit card action data to post goal and adaptive card related data.
    /// </summary>
    public class AdaptiveSubmitActionData
    {
        /// <summary>
        /// Gets or sets the Teams-specific action.
        /// </summary>
        [JsonProperty("msteams")]
        public CardAction MsTeams { get; set; }

        /// <summary>
        /// Gets or sets AAD object id of user who is saving personal goal details.
        /// </summary>
        [JsonProperty("UserAadObjectId")]
        public string UserAadObjectId { get; set; }

        /// <summary>
        /// Gets or sets activity id of card.
        /// </summary>
        [JsonProperty("ActivityId")]
        public string ActivityId { get; set; }

        /// <summary>
        /// Gets or sets adaptive card activity id of personal/team goal list card to refresh the same card after editing.
        /// </summary>
        [JsonProperty("ActivityCardId")]
        public string ActivityCardId { get; set; }

        /// <summary>
        /// Gets or sets adaptive action type.
        /// </summary>
        [JsonProperty("AdaptiveActionType")]
        public string AdaptiveActionType { get; set; }

        /// <summary>
        /// Gets or sets personal goal id from personal goal note detail.
        /// </summary>
        [JsonProperty("PersonalGoalId")]
        public string PersonalGoalId { get; set; }

        /// <summary>
        /// Gets or sets team id for aligning personal goal.
        /// </summary>
        [JsonProperty("TeamId")]
        public string TeamId { get; set; }

        /// <summary>
        /// Gets or sets goal note id to refresh the summary card.
        /// </summary>
        [JsonProperty("GoalNoteId")]
        public string GoalNoteId { get; set; }

        /// <summary>
        /// Gets or sets personal goal note id from personal goal note detail.
        /// </summary>
        [JsonProperty("PersonalGoalNoteId")]
        public string PersonalGoalNoteId { get; set; }

        /// <summary>
        /// Gets or sets goal cycle id, unique value for each goal cycle. This value is binded to team/personal goal
        /// list card to identify that goal cycle is ended or not.
        /// </summary>
        [JsonProperty("GoalCycleId")]
        public string GoalCycleId { get; set; }

        /// <summary>
        /// Gets or sets personal goal details received from task module submit action.
        /// </summary>
        [JsonProperty("PersonalGoalDetails")]
        public IEnumerable<PersonalGoalDetail> PersonalGoalDetails { get; set; }

        /// <summary>
        /// Gets or sets team goal details received from task module submit action.
        /// </summary>
        [JsonProperty("TeamGoalDetails")]
        public IEnumerable<TeamGoalDetail> TeamGoalDetails { get; set; }
    }
}
