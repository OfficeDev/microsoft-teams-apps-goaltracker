// <copyright file="HelpCard.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Cards
{
    using System.Collections.Generic;
    using AdaptiveCards;
    using Microsoft.Bot.Schema;
    using Microsoft.Extensions.Localization;
    using Microsoft.Teams.Apps.GoalTracker.Common;

    /// <summary>
    /// Implements help card.
    /// </summary>
    public static class HelpCard
    {
        /// <summary>
        /// Gets the help card for personal scope.
        /// </summary>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns>Help card attachment.</returns>
        public static Attachment GetHelpCardInPersonalChat(IStringLocalizer<Strings> localizer)
        {
            AdaptiveCard addNoteCard = new AdaptiveCard(Constants.AdaptiveCardVersion)
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveTextBlock
                    {
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Text = localizer.GetString("HelpCardContent"),
                        Wrap = true,
                    },
                    new AdaptiveTextBlock
                    {
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Text = localizer.GetString("HelpCardSetGoalsBulletPoint"),
                        Spacing = AdaptiveSpacing.Small,
                    },
                    new AdaptiveTextBlock
                    {
                        Text = localizer.GetString("HelpCardAddNotesBulletPoint"),
                        Spacing = AdaptiveSpacing.None,
                    },
                },
            };

            return new Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = addNoteCard,
            };
        }

        /// <summary>
        /// Gets the help card for team scope.
        /// </summary>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns>Help card attachment.</returns>
        public static Attachment GetHelpCardInChannel(IStringLocalizer<Strings> localizer)
        {
            AdaptiveCard addNoteCard = new AdaptiveCard(Constants.AdaptiveCardVersion)
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveTextBlock
                    {
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Text = localizer.GetString("HelpCardContent"),
                        Wrap = true,
                    },
                    new AdaptiveTextBlock
                    {
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Text = localizer.GetString("HelpCardSetTeamGoalsBulletPoint"),
                        Spacing = AdaptiveSpacing.Small,
                    },
                    new AdaptiveTextBlock
                    {
                        Text = localizer.GetString("HelpCardGoalStatusBulletPoint"),
                        Spacing = AdaptiveSpacing.None,
                    },
                },
            };

            return new Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = addNoteCard,
            };
        }
    }
}