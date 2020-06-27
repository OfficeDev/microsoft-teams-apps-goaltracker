// <copyright file="GoalStatusCard.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Cards
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using AdaptiveCards;
    using Microsoft.Bot.Schema;
    using Microsoft.Extensions.Localization;
    using Microsoft.Teams.Apps.GoalTracker.Common;
    using Microsoft.Teams.Apps.GoalTracker.Helpers;
    using Microsoft.Teams.Apps.GoalTracker.Models;

    /// <summary>
    ///  This class to create/send goal status card in Team Scope.
    /// </summary>
    public static class GoalStatusCard
    {
        /// <summary>
        /// Get the goal status card in channel.
        /// </summary>
        /// <param name="teamGoalStatuses">Team goal status for each team goal.</param>
        /// <param name="teamGoalStartDate">Team goal start date for team goal cycle.</param>
        /// <param name="teamGoalEndDate">Team goal end date for team goal cycle.</param>
        /// <param name="applicationBasePath">Application base URL.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns>Goal status card as attachment in teams channel.</returns>
        public static Attachment GetGoalStatusCard(IEnumerable<TeamGoalStatus> teamGoalStatuses, string teamGoalStartDate, string teamGoalEndDate, string applicationBasePath, IStringLocalizer<Strings> localizer)
        {
            var startDate = CardHelper.FormatDateStringToAdaptiveCardDateFormat(teamGoalStartDate);
            var endDate = CardHelper.FormatDateStringToAdaptiveCardDateFormat(teamGoalEndDate);
            AdaptiveCard goalStatusCard = new AdaptiveCard(Constants.AdaptiveCardVersion)
            {
                Height = AdaptiveHeight.Stretch,
                PixelMinHeight = 150,
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
                                    new AdaptiveTextBlock
                                    {
                                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                                        Text = localizer.GetString("GoalStatusCardTitle"),
                                        Weight = AdaptiveTextWeight.Bolder,
                                        Size = AdaptiveTextSize.Large,
                                        Wrap = true,
                                    },
                                },
                            },
                        },
                    },
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
                                    new AdaptiveTextBlock
                                    {
                                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                                        Text = localizer.GetString("GoalStatusGoalCycleText"),
                                        Wrap = true,
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
                                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                                        Text = $"{startDate} - {endDate}",
                                        Wrap = true,
                                        Weight = AdaptiveTextWeight.Bolder,
                                    },
                                },
                            },
                        },
                    },
                    new AdaptiveColumnSet
                    {
                        Columns = new List<AdaptiveColumn>
                        {
                            new AdaptiveColumn
                            {
                                Width = AdaptiveColumnWidth.Stretch,
                                Items = new List<AdaptiveElement>
                                { // Represents empty column required for alignment.
                                },
                            },
                        },
                    },
                },
            };

            List<AdaptiveElement> teamGoalStatusList = GetAllTeamGoalStatuses(teamGoalStatuses, applicationBasePath, localizer);
            goalStatusCard.Body.AddRange(teamGoalStatusList);

            var card = new Attachment()
            {
                ContentType = AdaptiveCard.ContentType,
                Content = goalStatusCard,
            };

            return card;
        }

        /// <summary>
        /// Get team goal status details card for all members for viewing status details.
        /// </summary>
        /// <param name="teamGoalStatuses">Team goal status for each team goal.</param>
        /// <param name="applicationBasePath">Application base URL.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns>Returns all team goal status card as attachment.</returns>
        public static List<AdaptiveElement> GetAllTeamGoalStatuses(IEnumerable<TeamGoalStatus> teamGoalStatuses, string applicationBasePath, IStringLocalizer<Strings> localizer)
        {
            teamGoalStatuses = teamGoalStatuses ?? throw new ArgumentNullException(nameof(teamGoalStatuses));
            int id = 1;
            List<AdaptiveElement> goalStatusList = new List<AdaptiveElement>();
            foreach (var teamGoalStatus in teamGoalStatuses)
            {
                goalStatusList.AddRange(GetIndividualTeamGoalStatus(teamGoalStatus, applicationBasePath, id, localizer));
                id += 1;
            }

            return goalStatusList;
        }

        /// <summary>
        /// Get user team goal status details card for viewing all user status details.
        /// </summary>
        /// <param name="teamGoalStatus">Team goal status for each team goal.</param>
        /// <param name="applicationBasePath">Application base URL.</param>
        /// <param name="id">Unique id for each row.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns>Returns individual goal card as attachment.</returns>
        public static List<AdaptiveElement> GetIndividualTeamGoalStatus(TeamGoalStatus teamGoalStatus, string applicationBasePath, int id, IStringLocalizer<Strings> localizer)
        {
            teamGoalStatus = teamGoalStatus ?? throw new ArgumentNullException(nameof(teamGoalStatus));
            var notStartedStatusCount = teamGoalStatus.NotStartedGoalCount == null ? 0 : teamGoalStatus.NotStartedGoalCount;
            var inProgressStatusCount = teamGoalStatus.InProgressGoalCount == null ? 0 : teamGoalStatus.InProgressGoalCount;
            var completedStatusCount = teamGoalStatus.CompletedGoalCount == null ? 0 : teamGoalStatus.CompletedGoalCount;
            var totalStatusCount = notStartedStatusCount + inProgressStatusCount + completedStatusCount;
            string cardContent = $"CardContent{id}";
            string chevronUp = $"ChevronUp{id}";
            string chevronDown = $"ChevronDown{id}";
            List<AdaptiveElement> individualTeamGoalStatus = new List<AdaptiveElement>
            {
                new AdaptiveColumnSet
                {
                    Spacing = AdaptiveSpacing.Medium,
                    Columns = new List<AdaptiveColumn>
                    {
                        new AdaptiveColumn
                        {
                            VerticalContentAlignment = AdaptiveVerticalContentAlignment.Center,
                            Width = AdaptiveColumnWidth.Stretch,
                            Items = new List<AdaptiveElement>
                            {
                                new AdaptiveTextBlock
                                {
                                    HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                                    Text = string.IsNullOrEmpty(teamGoalStatus.TeamGoalName) ? localizer.GetString("GoalStatusDeletedText") : teamGoalStatus.TeamGoalName,
                                    Wrap = true,
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
                                    HorizontalAlignment = AdaptiveHorizontalAlignment.Right,
                                    Spacing = AdaptiveSpacing.None,
                                    Text = string.Format(CultureInfo.InvariantCulture, localizer.GetString("GoalStatusAlignedText"), totalStatusCount),
                                    Wrap = true,
                                },
                            },
                        },
                        new AdaptiveColumn
                        {
                            Id = chevronDown,
                            VerticalContentAlignment = AdaptiveVerticalContentAlignment.Center,
                            Width = AdaptiveColumnWidth.Auto,
                            Items = new List<AdaptiveElement>
                            {
                                new AdaptiveImage
                                {
                                    AltText = localizer.GetString("GoalStatusChevronDownImageAltText"),
                                    PixelWidth = 16,
                                    PixelHeight = 8,
                                    SelectAction = new AdaptiveToggleVisibilityAction
                                    {
                                        Title = localizer.GetString("GoalStatusChevronDownImageAltText"),
                                        Type = "Action.ToggleVisibility",
                                        TargetElements = new List<AdaptiveTargetElement>
                                        {
                                            cardContent,
                                            chevronUp,
                                            chevronDown,
                                        },
                                    },
                                    Style = AdaptiveImageStyle.Default,
                                    Url = new Uri(applicationBasePath + "/Artifacts/chevronDown.png"),
                                    HorizontalAlignment = AdaptiveHorizontalAlignment.Right,
                                },
                            },
                        },
                        new AdaptiveColumn
                        {
                            Id = chevronUp,
                            IsVisible = false,
                            VerticalContentAlignment = AdaptiveVerticalContentAlignment.Center,
                            Width = AdaptiveColumnWidth.Auto,
                            Items = new List<AdaptiveElement>
                            {
                                new AdaptiveImage
                                {
                                    AltText = localizer.GetString("GoalStatusChevronUpImageAltText"),
                                    PixelWidth = 16,
                                    PixelHeight = 8,
                                    SelectAction = new AdaptiveToggleVisibilityAction
                                    {
                                        Title = localizer.GetString("GoalStatusChevronUpImageAltText"),
                                        Type = "Action.ToggleVisibility",
                                        TargetElements = new List<AdaptiveTargetElement>
                                        {
                                            cardContent,
                                            chevronUp,
                                            chevronDown,
                                        },
                                    },
                                    Style = AdaptiveImageStyle.Default,
                                    Url = new Uri(applicationBasePath + "/Artifacts/chevronUp.png"),
                                    HorizontalAlignment = AdaptiveHorizontalAlignment.Right,
                                },
                            },
                        },
                    },
                },
                new AdaptiveContainer
                {
                    Id = cardContent,
                    IsVisible = false,
                    Style = AdaptiveContainerStyle.Emphasis,
                    Items = new List<AdaptiveElement>
                    {
                        new AdaptiveColumnSet
                        {
                            Spacing = AdaptiveSpacing.None,
                            Columns = new List<AdaptiveColumn>
                            {
                                new AdaptiveColumn
                                {
                                    VerticalContentAlignment = AdaptiveVerticalContentAlignment.Center,
                                    Width = AdaptiveColumnWidth.Stretch,
                                    Items = new List<AdaptiveElement>
                                    {
                                        new AdaptiveTextBlock
                                        {
                                            HorizontalAlignment = AdaptiveHorizontalAlignment.Center,
                                            Spacing = AdaptiveSpacing.None,
                                            Text = localizer.GetString("GoalStatusNotStartedHeader"),
                                            Wrap = true,
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
                                            HorizontalAlignment = AdaptiveHorizontalAlignment.Center,
                                            Spacing = AdaptiveSpacing.None,
                                            Text = localizer.GetString("GoalStatusInProgressHeader"),
                                            Wrap = true,
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
                                            HorizontalAlignment = AdaptiveHorizontalAlignment.Center,
                                            Spacing = AdaptiveSpacing.None,
                                            Text = localizer.GetString("GoalStatusCompletedHeader"),
                                            Wrap = true,
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
                                            HorizontalAlignment = AdaptiveHorizontalAlignment.Center,
                                            Spacing = AdaptiveSpacing.None,
                                            Text = localizer.GetString("GoalStatusTotalHeader"),
                                            Wrap = true,
                                        },
                                    },
                                },
                            },
                        },
                        new AdaptiveColumnSet
                        {
                            Spacing = AdaptiveSpacing.None,
                            Columns = new List<AdaptiveColumn>
                            {
                                new AdaptiveColumn
                                {
                                    VerticalContentAlignment = AdaptiveVerticalContentAlignment.Center,
                                    Width = AdaptiveColumnWidth.Stretch,
                                    Items = new List<AdaptiveElement>
                                    {
                                        new AdaptiveTextBlock
                                        {
                                            HorizontalAlignment = AdaptiveHorizontalAlignment.Center,
                                            Spacing = AdaptiveSpacing.None,
                                            Text = notStartedStatusCount.ToString(),
                                            Wrap = true,
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
                                            HorizontalAlignment = AdaptiveHorizontalAlignment.Center,
                                            Spacing = AdaptiveSpacing.None,
                                            Text = inProgressStatusCount.ToString(),
                                            Wrap = true,
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
                                            HorizontalAlignment = AdaptiveHorizontalAlignment.Center,
                                            Spacing = AdaptiveSpacing.None,
                                            Text = completedStatusCount.ToString(),
                                            Wrap = true,
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
                                            HorizontalAlignment = AdaptiveHorizontalAlignment.Center,
                                            Spacing = AdaptiveSpacing.None,
                                            Text = totalStatusCount.ToString(),
                                            Wrap = true,
                                        },
                                    },
                                },
                            },
                        },
                    },
                },
            };

            return individualTeamGoalStatus;
        }
    }
}