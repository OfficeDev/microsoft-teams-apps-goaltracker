// <copyright file="PersonalGoalDetail.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Models
{
    using System.ComponentModel.DataAnnotations;
    using Microsoft.Azure.Search;
    using Microsoft.WindowsAzure.Storage.Table;
    using Newtonsoft.Json;

    /// <summary>
    /// Class containing personal goal details which are updated by individual user.
    /// </summary>
    public class PersonalGoalDetail : TableEntity
    {
        /// <summary>
        /// Gets or sets Azure Active Directory object id of the user.
        /// </summary>
        [Required]
        [JsonProperty("UserAadObjectId")]
        public string UserAadObjectId
        {
            get { return this.PartitionKey; }
            set { this.PartitionKey = value; }
        }

        /// <summary>
        /// Gets or sets unique id of each personal goal.
        /// </summary>
        [IsSearchable]
        [IsFilterable]
        [Key]
        [Required]
        [JsonProperty("PersonalGoalId")]
        public string PersonalGoalId
        {
            get { return this.RowKey; }
            set { this.RowKey = value; }
        }

        /// <summary>
        /// Gets or sets date at which personal goals are created.
        /// </summary>
        [Required]
        [JsonProperty("CreatedOn")]
        public string CreatedOn { get; set; }

        /// <summary>
        /// Gets or sets name of the user who created the personal goals.
        /// </summary>
        [Required]
        [JsonProperty("CreatedBy")]
        public string CreatedBy { get; set; }

        /// <summary>
        /// Gets or sets date at which personal goals are updated.
        /// </summary>
        [JsonProperty("LastModifiedOn")]
        public string LastModifiedOn { get; set; }

        /// <summary>
        /// Gets or sets name of user who modified the personal goals.
        /// </summary>
        [JsonProperty("LastModifiedBy")]
        public string LastModifiedBy { get; set; }

        /// <summary>
        /// Gets or sets activity id of the adaptive card.
        /// </summary>
        [JsonProperty("AdaptiveCardActivityId")]
        public string AdaptiveCardActivityId { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether personal goal cycle is active or not.
        /// </summary>
        [IsFilterable]
        [JsonProperty("IsActive")]
        public bool IsActive { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether personal goal is aligned with team goal or not.
        /// </summary>
        [JsonProperty("IsAligned")]
        public bool IsAligned { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether personal goal is deleted or not.
        /// </summary>
        [IsFilterable]
        [JsonProperty("IsDeleted")]
        public bool IsDeleted { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether user wants to receive reminders.
        /// </summary>
        [JsonProperty("IsReminderActive")]
        public bool IsReminderActive { get; set; }

        /// <summary>
        /// Gets or sets the conversation id of the personal chat with bot on 1:1.
        /// </summary>
        [JsonProperty("ConversationId")]
        public string ConversationId { get; set; }

        /// <summary>
        /// Gets or sets description of each personal goal.
        /// </summary>
        [Required]
        [MaxLength(300)]
        [JsonProperty("GoalName")]
        public string GoalName { get; set; }

        /// <summary>
        /// Gets or sets status of the each personal goal i.e. not started/in progress/completed.
        /// </summary>
        [IsFacetable]
        [IsSortable]
        [IsFilterable]
        [Range(0, 2)]
        [JsonProperty("Status")]
        public int Status { get; set; }

        /// <summary>
        /// Gets or sets start date of the personal goals.
        /// </summary>
        [Required]
        [JsonProperty("StartDate")]
        public string StartDate { get; set; }

        /// <summary>
        /// Gets or sets end date of the personal goals.
        /// </summary>
        [Required]
        [JsonProperty("EndDate")]
        public string EndDate { get; set; }

        /// <summary>
        /// Gets or sets reminder frequency of the personal goals i.e. Weekly/Bi-weekly/Monthly/Quarterly.
        /// </summary>
        [Range(0, 3)]
        [JsonProperty("ReminderFrequency")]
        public int ReminderFrequency { get; set; }

        /// <summary>
        /// Gets or sets goal cycle id, unique value for each goal cycle.
        /// This value is binded to personal goal list card to identify that goal cycle is ended or not.
        /// </summary>
        [Required]
        [JsonProperty("GoalCycleId")]
        public string GoalCycleId { get; set; }

        /// <summary>
        /// Gets or sets team id to identify in which team user has aligned his/her personal goals.
        /// </summary>
        [IsSearchable]
        [IsFilterable]
        [JsonProperty("TeamId")]
        public string TeamId { get; set; }

        /// <summary>
        /// Gets or sets unique id of each team goal to identify personal goal is aligned with.
        /// </summary>
        [IsFacetable]
        [IsFilterable]
        [JsonProperty("TeamGoalId")]
        public string TeamGoalId { get; set; }

        /// <summary>
        /// Gets or sets personal goal end date in UTC format to fetch the records from storage for sending goal reminder to personal bot.
        /// </summary>
        [JsonProperty("EndDateUTC")]
        public string EndDateUTC { get; set; }

        /// <summary>
        /// Gets or sets service endpoint where operations concerning the referenced conversation.
        /// </summary>
        [JsonProperty("ServiceURL")]
        public string ServiceUrl { get; set; }
    }
}
