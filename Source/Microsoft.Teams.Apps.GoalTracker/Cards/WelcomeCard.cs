// <copyright file="WelcomeCard.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Cards
{
    using System;
    using System.Collections.Generic;
    using AdaptiveCards;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Localization;
    using Microsoft.Teams.Apps.GoalTracker.Common;
    using Microsoft.Teams.Apps.GoalTracker.Models;
    using Newtonsoft.Json;

    /// <summary>
    ///  Class having method to return welcome card attachment for personal chat and channel.
    /// </summary>
    public static class WelcomeCard
    {
        /// <summary>
        /// Represent welcome card icon width.
        /// </summary>
        private const uint WelcomeCardIconWidth = 56;

        /// <summary>
        /// Represent welcome card icon height.
        /// </summary>
        private const uint WelcomeCardIconHeight = 56;

        /// <summary>
        /// Get welcome card attachment to show on Microsoft Teams personal scope when bot is installed personally.
        /// </summary>
        /// <param name="applicationBasePath">Application base URL to get the logo of the application.</param>
        /// <param name="localizer">The current cultures' string localizer</param>
        /// <returns>User welcome card.</returns>
        public static Attachment GetWelcomeCardAttachmentForPersonalChat(string applicationBasePath, IStringLocalizer<Strings> localizer)
        {
            AdaptiveCard welcomeCard = new AdaptiveCard(new AdaptiveSchemaVersion(Constants.AdaptiveCardVersion))
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveColumnSet
                    {
                        Columns = new List<AdaptiveColumn>
                        {
                            new AdaptiveColumn
                            {
                                VerticalContentAlignment = AdaptiveVerticalContentAlignment.Center,
                                Width = AdaptiveColumnWidth.Auto,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveImage
                                    {
                                        Url = new Uri($"{applicationBasePath}/Artifacts/appLogo.png"),
                                        Size = AdaptiveImageSize.Large,
                                        AltText = localizer.GetString("AltTextForWelcomeCardImage"),
                                        HorizontalAlignment = AdaptiveHorizontalAlignment.Center,
                                        PixelHeight = WelcomeCardIconHeight,
                                        PixelWidth = WelcomeCardIconWidth,
                                    },
                                },
                            },
                            new AdaptiveColumn
                            {
                                VerticalContentAlignment = AdaptiveVerticalContentAlignment.Center,
                                Width = AdaptiveColumnWidth.Stretch,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Text = localizer.GetString("WelcomeCardTitle"),
                                        Spacing = AdaptiveSpacing.None,
                                        Weight = AdaptiveTextWeight.Bolder,
                                        Wrap = true,
                                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                                    },
                                    new AdaptiveTextBlock
                                    {
                                        Text = localizer.GetString("WelcomeCardContent"),
                                        Wrap = true,
                                        Spacing = AdaptiveSpacing.None,
                                        IsSubtle = true,
                                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                                    },
                                },
                            },
                        },
                    },
                    new AdaptiveTextBlock
                    {
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Text = localizer.GetString("WelcomeSubHeaderText"),
                        Spacing = AdaptiveSpacing.Small,
                        Wrap = true,
                    },
                    new AdaptiveTextBlock
                    {
                        Text = localizer.GetString("SetGoalsBulletPoint"),
                        Spacing = AdaptiveSpacing.None,
                        Wrap = true,
                    },
                    new AdaptiveTextBlock
                    {
                        Text = localizer.GetString("SaveNotesBulletPoint"),
                        Spacing = AdaptiveSpacing.None,
                        Wrap = true,
                    },
                    new AdaptiveTextBlock
                    {
                        Text = localizer.GetString("ContentText"),
                        Spacing = AdaptiveSpacing.Small,
                    },
                },
                Actions = new List<AdaptiveAction>
                {
                    new AdaptiveSubmitAction
                    {
                        Title = localizer.GetString("SetGoalsButtonText"),
                        Data = new AdaptiveSubmitActionData
                        {
                            MsTeams = new TaskModuleAction(localizer.GetString("SetGoalsButtonText"), new { data = JsonConvert.SerializeObject(new AdaptiveSubmitActionData { AdaptiveActionType = Constants.SetPersonalGoalsCommand }) }),
                        },
                    },
                },
            };

            return new Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = welcomeCard,
            };
        }

        /// <summary>
        /// Get welcome card attachment to show on Microsoft Teams team scope when bot is installed in team.
        /// </summary>
        /// <param name="localizer">The current cultures' string localizer</param>
        /// <returns>Teams welcome card.</returns>
        public static Attachment GetWelcomeCardAttachmentForChannel(IStringLocalizer<Strings> localizer)
        {
            AdaptiveCard welcomeCard = new AdaptiveCard(new AdaptiveSchemaVersion(Constants.AdaptiveCardVersion))
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveTextBlock
                    {
                        Text = localizer.GetString("TeamsWelcomeCardContent"),
                        Wrap = true,
                        Spacing = AdaptiveSpacing.None,
                    },
                    new AdaptiveTextBlock
                    {
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Text = localizer.GetString("TeamsWelcomeSubHeaderText"),
                        Spacing = AdaptiveSpacing.None,
                    },
                    new AdaptiveTextBlock
                    {
                        Text = localizer.GetString("SetTeamGoalsBulletPoint"),
                        Wrap = true,
                        Spacing = AdaptiveSpacing.None,
                    },
                    new AdaptiveTextBlock
                    {
                        Text = localizer.GetString("TeamGoalsBulletPoint"),
                        Wrap = true,

                        Spacing = AdaptiveSpacing.None,
                    },
                    new AdaptiveTextBlock
                    {
                        Text = localizer.GetString("AlignTeamGoalsBulletPoint"),
                        Wrap = true,
                        Spacing = AdaptiveSpacing.None,
                    },
                    new AdaptiveTextBlock
                    {
                        Text = localizer.GetString("ContentText"),
                        Spacing = AdaptiveSpacing.Small,
                    },
                },
                Actions = new List<AdaptiveAction>
                {
                    new AdaptiveSubmitAction
                    {
                        Title = localizer.GetString("SetTeamGoalsButtonText"),
                        Data = new AdaptiveSubmitActionData
                        {
                            MsTeams = new TaskModuleAction(localizer.GetString("SetTeamGoalsButtonText"), new { data = JsonConvert.SerializeObject(new AdaptiveSubmitActionData { AdaptiveActionType = Constants.SetTeamGoalsCommand }) }),
                        },
                    },
                },
            };

            return new Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = welcomeCard,
            };
        }
    }
}
