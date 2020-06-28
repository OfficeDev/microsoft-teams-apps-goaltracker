// <copyright file="Constants.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Common
{
    /// <summary>
    /// Constant values that are used in multiple files.
    /// </summary>
    public static class Constants
    {
        /// <summary>
        /// Command to send adaptive card with set goals button in personal bot.
        /// </summary>
        public const string SetGoalsCommand = "SET GOALS";

        /// <summary>
        /// Command to send adaptive card with set goals button in personal bot.
        /// </summary>
        public const string EditGoalsCommand = "EDIT GOALS";

        /// <summary>
        /// Command text to fetch set personal goals task module.
        /// </summary>
        public const string SetPersonalGoalsCommand = "SET PERSONAL GOALS";

        /// <summary>
        /// Command text to fetch edit personal goals task module.
        /// </summary>
        public const string EditPersonalGoalsCommand = "EDIT PERSONAL GOALS";

        /// <summary>
        /// Command text to fetch set team goals task module or send adaptive card with set team goals button in team.
        /// </summary>
        public const string SetTeamGoalsCommand = "SET TEAM GOALS";

        /// <summary>
        /// Command text to fetch edit team goals task module.
        /// </summary>
        public const string EditTeamGoalsCommand = "EDIT TEAM GOALS";

        /// <summary>
        /// Command to fetch add note task module or send adaptive card with add note button.
        /// </summary>
        public const string AddNoteCommand = "ADD NOTE";

        /// <summary>
        /// Command to fetch edit note task module.
        /// </summary>
        public const string EditNoteCommand = "EDIT NOTE";

        /// <summary>
        /// Command text to fetch align goal task module.
        /// </summary>
        public const string AlignGoalCommand = "ALIGN GOAL";

        /// <summary>
        /// Command to send goal status adaptive card to know the status of team goals.
        /// </summary>
        public const string GoalStatusCommand = "GOAL STATUS";

        /// <summary>
        /// Maximum number of notes that can be added to a goal.
        /// </summary>
        public const int MaximumNumberOfNotes = 10;

        /// <summary>
        /// Maximum number of goals that can be added by team owner or user.
        /// </summary>
        public const int MaximumNumberOfGoals = 15;

        /// <summary>
        /// Described adaptive card version to be used. Version can be upgraded or changed using this value.
        /// </summary>
        public const string AdaptiveCardVersion = "1.2";

        /// <summary>
        /// Date time format to support adaptive card text feature.
        /// </summary>
        /// <remarks>
        /// refer adaptive card text feature https://docs.microsoft.com/en-us/adaptive-cards/authoring-cards/text-features#datetime-formatting-and-localization.
        /// </remarks>
        public const string Rfc3339DateTimeFormat = "yyyy'-'MM'-'dd'T'HH':'mm':'ss'Z'";

        /// <summary>
        /// Date time format to show in list card.
        /// </summary>
        public const string ListCardDateTimeFormat = "d";

        /// <summary>
        /// Date time format to compare with end date UTC for sending reminder or to end goal cycle.
        /// </summary>
        public const string UTCDateFormat = "MM-dd-yyyy";

        /// <summary>
        /// Message back card action.
        /// </summary>
        public const string MessageBackActionType = "messageBack";

        /// <summary>
        /// Represents task module task/fetch string.
        /// </summary>
        public const string TaskModuleFetchType = "task/fetch";

        /// <summary>
        /// Represents task module task/submit string.
        /// </summary>
        public const string TaskModuleSubmitType = "task/submit";

        /// <summary>
        /// Represents personal conversation type.
        /// </summary>
        public const string PersonalConversationType = "personal";

        /// <summary>
        /// Represents channel conversation type.
        /// </summary>
        public const string ChannelConversationType = "channel";

        /// <summary>
        /// Represents channel conversation id.
        /// </summary>
        public const string TeamsBotFrameworkChannelId = "msteams";
    }
}
