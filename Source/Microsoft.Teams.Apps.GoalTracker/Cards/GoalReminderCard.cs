// <copyright file="GoalReminderCard.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Cards
{
    using System;
    using System.Collections.Generic;
    using AdaptiveCards;
    using Microsoft.Bot.Schema;
    using Microsoft.Extensions.Localization;
    using Microsoft.Teams.Apps.GoalTracker.Common;

    /// <summary>
    /// Class contains goal reminder card.
    /// </summary>
    public static class GoalReminderCard
    {
        /// <summary>
        /// Gets the goal reminder card.
        /// </summary>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="applicationManifestId">Application manifest id.</param>
        /// <param name="goalsTabEntityId">Goals tab entity id for personal Bot.</param>
        /// <param name="reminderTypeString">Reminder type string based on reminder frequency.</param>
        /// <param name="reminderTypeStringColor">Reminder type string color based on reminder frequency.</param>
        /// <returns>Goal reminder card attachment.</returns>
        public static Attachment GetGoalReminderCard(IStringLocalizer<Strings> localizer, string applicationManifestId, string goalsTabEntityId, string reminderTypeString, AdaptiveTextColor reminderTypeStringColor)
        {
            AdaptiveCard goalReminderCard = new AdaptiveCard(Constants.AdaptiveCardVersion)
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveTextBlock
                    {
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Wrap = true,
                        Text = localizer.GetString("GoalReminderCardTitle"),
                        Weight = AdaptiveTextWeight.Bolder,
                        Size = AdaptiveTextSize.Large,
                    },
                    new AdaptiveTextBlock
                    {
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Wrap = true,
                        Text = reminderTypeString,
                        Color = reminderTypeStringColor,
                    },
                    new AdaptiveTextBlock
                    {
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Wrap = true,
                        Text = localizer.GetString("GoalReminderCardContentText"),
                    },
                },
            };

            goalReminderCard.Actions = new List<AdaptiveAction>
            {
                new AdaptiveOpenUrlAction
                {
                    Title = localizer.GetString("ViewGoalsButtonText"),
                    Url = new Uri($"https://teams.microsoft.com/l/entity/{applicationManifestId}/{goalsTabEntityId}"), // Open Goals tab (deep link).
                },
            };

            return new Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = goalReminderCard,
            };
        }
    }
}
