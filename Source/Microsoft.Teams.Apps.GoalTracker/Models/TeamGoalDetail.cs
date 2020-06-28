// <copyright file="TeamGoalDetail.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Models
{
    using System.ComponentModel.DataAnnotations;
    using Microsoft.WindowsAzure.Storage.Table;
    using Newtonsoft.Json;

    /// <summary>
    /// Class containing team goal details which are updated by team owner.
    /// </summary>
    public class TeamGoalDetail : TableEntity
    {
        /// <summary>
        /// Gets or sets team id.
        /// </summary>
        [Required]
        [JsonProperty("TeamId")]
        public string TeamId
        {
            get { return this.PartitionKey; }
            set { this.PartitionKey = value; }
        }

        /// <summary>
        /// Gets or sets unique id of each team goal.
        /// </summary>
        [Required]
        [JsonProperty("TeamGoalId")]
        public string TeamGoalId
        {
            get { return this.RowKey; }
            set { this.RowKey = value; }
        }

        /// <summary>
        /// Gets or sets date at which team goals are created.
        /// </summary>
        [Required]
        [JsonProperty("CreatedOn")]
        public string CreatedOn { get; set; }

        /// <summary>
        /// Gets or sets name of the user who created the team goals.
        /// </summary>
        [Required]
        [JsonProperty("CreatedBy")]
        public string CreatedBy { get; set; }

        /// <summary>
        /// Gets or sets date at which team goals are updated.
        /// </summary>
        [JsonProperty("LastModifiedOn")]
        public string LastModifiedOn { get; set; }

        /// <summary>
        /// Gets or sets name of user who modified the team goals.
        /// </summary>
        [JsonProperty("LastModifiedBy")]
        public string LastModifiedBy { get; set; }

        /// <summary>
        /// Gets or sets activity id of the adaptive card.
        /// </summary>
        [JsonProperty("AdaptiveCardActivityId")]
        public string AdaptiveCardActivityId { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether team goal cycle is active or not.
        /// </summary>
        [JsonProperty("IsActive")]
        public bool IsActive { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether team goal is deleted or not.
        /// </summary>
        [JsonProperty("IsDeleted")]
        public bool IsDeleted { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether reminder needs to be sent to team and team members.
        /// </summary>
        [JsonProperty("IsReminderActive")]
        public bool IsReminderActive { get; set; }

        /// <summary>
        /// Gets or sets the conversation id of the channel.
        /// </summary>
        [JsonProperty("ChannelConversationId")]
        public string ChannelConversationId { get; set; }

        /// <summary>
        /// Gets or sets reminder frequency of the team goals i.e. Weekly/Bi-weekly/Monthly/Quarterly.
        /// </summary>
        [Range(0, 3)]
        [JsonProperty("ReminderFrequency")]
        public int ReminderFrequency { get; set; }

        /// <summary>
        /// Gets or sets description of each team goal.
        /// </summary>
        [Required]
        [JsonProperty("TeamGoalName")]
        public string TeamGoalName { get; set; }

        /// <summary>
        /// Gets or sets start date of the team goals.
        /// </summary>
        [Required]
        [JsonProperty("TeamGoalStartDate")]
        public string TeamGoalStartDate { get; set; }

        /// <summary>
        /// Gets or sets end date of the team goals.
        /// </summary>
        [Required]
        [JsonProperty("TeamGoalEndDate")]
        public string TeamGoalEndDate { get; set; }

        /// <summary>
        /// Gets or sets team goal end date in UTC format to fetch the records from storage for sending goal reminder in team.
        /// </summary>
        [JsonProperty("TeamGoalEndDateUTC")]
        public string TeamGoalEndDateUTC { get; set; }

        /// <summary>
        /// Gets or sets activity service URL.
        /// </summary>
        [JsonProperty("ServiceURL")]
        public string ServiceUrl { get; set; }

        /// <summary>
        /// Gets or sets goal cycle id, unique value for each goal cycle.
        /// This value is binded to team goal list card sent in team and to team members to identify that goal cycle is ended or not.
        /// </summary>
        [Required]
        [JsonProperty("GoalCycleId")]
        public string GoalCycleId { get; set; }
    }
}
