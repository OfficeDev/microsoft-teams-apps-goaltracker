// <copyright file="PersonalGoalNoteDetail.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Models
{
    using System.ComponentModel.DataAnnotations;
    using Microsoft.WindowsAzure.Storage.Table;
    using Newtonsoft.Json;

    /// <summary>
    /// Class containing personal note details for each personal goal.
    /// </summary>
    public class PersonalGoalNoteDetail : TableEntity
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
        /// Gets or sets unique id of each personal goal note.
        /// </summary>
        [Required]
        [JsonProperty("PersonalGoalNoteId")]
        public string PersonalGoalNoteId
        {
            get { return this.RowKey; }
            set { this.RowKey = value; }
        }

        /// <summary>
        /// Gets or sets unique id of each personal goal.
        /// </summary>
        [Required]
        [JsonProperty("PersonalGoalId")]
        public string PersonalGoalId { get; set; }

        /// <summary>
        /// Gets or sets date at which personal note is created.
        /// </summary>
        [Required]
        [JsonProperty("CreatedOn")]
        public string CreatedOn { get; set; }

        /// <summary>
        /// Gets or sets name of the user who created the personal note.
        /// </summary>
        [Required]
        [JsonProperty("CreatedBy")]
        public string CreatedBy { get; set; }

        /// <summary>
        /// Gets or sets date at which personal note is updated.
        /// </summary>
        [JsonProperty("LastModifiedOn")]
        public string LastModifiedOn { get; set; }

        /// <summary>
        /// Gets or sets name of user who modified the personal note.
        /// </summary>
        [JsonProperty("LastModifiedBy")]
        public string LastModifiedBy { get; set; }

        /// <summary>
        /// Gets or sets activity id of the adaptive card.
        /// </summary>
        [JsonProperty("AdaptiveCardActivityId")]
        public string AdaptiveCardActivityId { get; set; }

        /// <summary>
        /// Gets or sets the conversation id of the personal chat with bot on 1:1.
        /// </summary>
        [JsonProperty("ConversationId")]
        public string ConversationId { get; set; }

        /// <summary>
        /// Gets or sets note description for individual personal goal.
        /// </summary>
        [Required]
        [MaxLength(1000)]
        [JsonProperty("PersonalGoalNoteDescription")]
        public string PersonalGoalNoteDescription { get; set; }

        /// <summary>
        /// Gets or sets source name of user from which note is shared. If add note task module invoked from message action then source name will be
        /// name of user whose message is referred from message action and will be non-editable.
        /// </summary>
        [MaxLength(100)]
        [JsonProperty("SourceName")]
        public string SourceName { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether personal goal cycle is active or not.
        /// </summary>
        [JsonProperty("IsActive")]
        public bool IsActive { get; set; }
    }
}
