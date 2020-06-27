// <copyright file="GoalCard.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Cards
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using AdaptiveCards;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Localization;
    using Microsoft.Teams.Apps.GoalTracker.Common;
    using Microsoft.Teams.Apps.GoalTracker.Models;
    using Newtonsoft.Json;

    /// <summary>
    /// Class having method to return personal and team goals related card attachment.
    /// </summary>
    public static class GoalCard
    {
        /// <summary>
        /// Gets the set goals card in personal scope.
        /// </summary>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns>Returns an attachment of set goals card in personal scope.</returns>
        public static Attachment GetSetPersonalGoalsCard(IStringLocalizer<Strings> localizer)
        {
            AdaptiveCard setGoalsCommandCard = new AdaptiveCard(Constants.AdaptiveCardVersion)
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveTextBlock
                    {
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Text = localizer.GetString("SetGoalsCommandCardText"),
                        Wrap = true,
                    },
                },
                Actions = new List<AdaptiveAction>
                {
                    new AdaptiveSubmitAction
                    {
                        Title = localizer.GetString("SetGoalsActionText"),
                        Data = new AdaptiveSubmitActionData
                        {
                            MsTeams = new TaskModuleAction(localizer.GetString("SetGoalsActionText"), new { data = JsonConvert.SerializeObject(new AdaptiveSubmitActionData { AdaptiveActionType = Constants.SetPersonalGoalsCommand }) }),
                        },
                    },
                },
            };

            return new Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = setGoalsCommandCard,
            };
        }

        /// <summary>
        /// Gets the edit goals card in personal scope.
        /// </summary>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns>Returns an attachment of edit goals card in personal scope.</returns>
        public static Attachment GetEditPersonalGoalsCard(IStringLocalizer<Strings> localizer)
        {
            AdaptiveCard editGoalsCommandCard = new AdaptiveCard(Constants.AdaptiveCardVersion)
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveTextBlock
                    {
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Text = localizer.GetString("EditGoalsCommandCardText"),
                        Wrap = true,
                    },
                },
                Actions = new List<AdaptiveAction>
                {
                    new AdaptiveSubmitAction
                    {
                        Title = localizer.GetString("EditGoalsActionText"),
                        Data = new AdaptiveSubmitActionData
                        {
                            MsTeams = new TaskModuleAction(localizer.GetString("EditGoalsActionText"), new { data = JsonConvert.SerializeObject(new AdaptiveSubmitActionData { AdaptiveActionType = Constants.EditPersonalGoalsCommand }) }),
                        },
                    },
                },
            };

            return new Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = editGoalsCommandCard,
            };
        }

        /// <summary>
        /// Gets the set goals card in team scope.
        /// </summary>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns>Returns an attachment of set team goals card in team scope.</returns>
        public static Attachment GetSetTeamGoalsCard(IStringLocalizer<Strings> localizer)
        {
            AdaptiveCard setTeamsGoalsCard = new AdaptiveCard(Constants.AdaptiveCardVersion)
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveTextBlock
                    {
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Text = localizer.GetString("SetGoalsCardTitleTextTeam"),
                        Wrap = true,
                    },
                    new AdaptiveTextBlock
                    {
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Text = localizer.GetString("SetTeamGoalsCardText"),
                        Wrap = true,
                        Weight = AdaptiveTextWeight.Bolder,
                        Size = AdaptiveTextSize.Large,
                    },
                    new AdaptiveTextBlock
                    {
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Text = localizer.GetString("SetTeamGoalsCardTextBody"),
                        Wrap = true,
                    },
                },
                Actions = new List<AdaptiveAction>
                {
                    new AdaptiveSubmitAction
                    {
                        Title = localizer.GetString("SetTeamGoalsActionText"),
                        Data = new AdaptiveSubmitActionData
                        {
                            MsTeams = new TaskModuleAction(localizer.GetString("SetTeamGoalsActionText"), new { data = JsonConvert.SerializeObject(new AdaptiveSubmitActionData { AdaptiveActionType = Constants.SetTeamGoalsCommand }) }),
                        },
                    },
                },
            };

            return new Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = setTeamsGoalsCard,
            };
        }

        /// <summary>
        /// Gets the edit goals card in team scope.
        /// </summary>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="teamId">Team id to edit team goals.</param>
        /// <returns>Returns an attachment of edit team goals card in team scope.</returns>
        public static Attachment GetEditTeamGoalsCard(IStringLocalizer<Strings> localizer, string teamId)
        {
            AdaptiveCard editTeamsGoalsCard = new AdaptiveCard(Constants.AdaptiveCardVersion)
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveTextBlock
                    {
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Text = localizer.GetString("EditGoalsCardTitleTextTeam"),
                        Wrap = true,
                    },
                },
                Actions = new List<AdaptiveAction>
                {
                    new AdaptiveSubmitAction
                    {
                        Title = localizer.GetString("EditTeamGoalsActionText"),
                        Data = new AdaptiveSubmitActionData
                        {
                            MsTeams = new TaskModuleAction(localizer.GetString("EditTeamGoalsActionText"), new { data = JsonConvert.SerializeObject(new AdaptiveSubmitActionData { AdaptiveActionType = Constants.EditTeamGoalsCommand, TeamId = teamId }) }),
                        },
                    },
                },
            };

            return new Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = editTeamsGoalsCard,
            };
        }

        /// <summary>
        /// Method to show teams goals details card in team.
        /// </summary>
        /// <param name="applicationBasePath">Application base URL of the application.</param>
        /// <param name="teamGoalDetails">Team goal values entered by user.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="teamGoalsCreator">Name of the user who created/edited team goals.</param>
        /// <param name="goalCycleId">Goal cycle id to identify that goal cycle is ended or not.</param>
        /// <param name="isCardForTeamMembers">Determines whether card is sent in personal or team scope.</param>
        /// <returns>Returns an attachment of teams goals details card in team.</returns>
        public static Attachment GetTeamGoalDetailsListCard(string applicationBasePath, IEnumerable<TeamGoalDetail> teamGoalDetails, IStringLocalizer<Strings> localizer, string teamGoalsCreator, string goalCycleId, bool isCardForTeamMembers = false)
        {
            teamGoalDetails = teamGoalDetails ?? throw new ArgumentNullException(nameof(teamGoalDetails));

            var teamId = teamGoalDetails.Select(goal => goal.TeamId).First();
            var teamGoalCycleStartDate = teamGoalDetails.Select(goal => goal.TeamGoalStartDate).First();
            teamGoalCycleStartDate = DateTime.Parse(teamGoalCycleStartDate, CultureInfo.InvariantCulture).ToUniversalTime()
                .ToString(Constants.ListCardDateTimeFormat, CultureInfo.InvariantCulture);

            var teamGoalCycleEndDate = teamGoalDetails.Select(goal => goal.TeamGoalEndDate).First();
            teamGoalCycleEndDate = DateTime.Parse(teamGoalCycleEndDate, CultureInfo.InvariantCulture).ToUniversalTime()
                .ToString(Constants.ListCardDateTimeFormat, CultureInfo.InvariantCulture);

            var reminder = teamGoalDetails.Select(goal => goal.ReminderFrequency).First();
            var reminderFrequency = (ReminderFrequency)reminder;
            ListCard teamGoalDetailsListCard = new ListCard
            {
                Title = isCardForTeamMembers
                        ? localizer.GetString("TeamGoalListCardTitleForTeamMembers")
                        : localizer.GetString("TeamGoalListCardTitleForTeam"),
            };

            teamGoalDetailsListCard.Items.Add(new ListItem
            {
                Title = isCardForTeamMembers
                        ? localizer.GetString("TeamGoalCardCycleTextForTeamMembers", teamGoalsCreator, teamGoalCycleStartDate, teamGoalCycleEndDate)
                        : localizer.GetString("TeamGoalCardCycleTextForTeam", teamGoalCycleStartDate, teamGoalCycleEndDate),
                Type = "section",
            });

            if (isCardForTeamMembers)
            {
                foreach (var teamGoalDetailEntity in teamGoalDetails)
                {
                    teamGoalDetailsListCard.Items.Add(new ListItem
                    {
                        Id = teamGoalDetailEntity.TeamGoalId,
                        Title = teamGoalDetailEntity.TeamGoalName,
                        Subtitle = reminderFrequency.ToString(),
                        Type = "resultItem",
                        Icon = $"{applicationBasePath}/Artifacts/listIcon.png",
                    });
                }
            }
            else
            {
                foreach (var teamGoalDetailEntity in teamGoalDetails)
                {
                    teamGoalDetailsListCard.Items.Add(new ListItem
                    {
                        Id = teamGoalDetailEntity.TeamGoalId,
                        Title = teamGoalDetailEntity.TeamGoalName,
                        Subtitle = reminderFrequency.ToString(),
                        Type = "resultItem",
                        Tap = new TaskModuleAction(localizer.GetString("EditButtonText"), new AdaptiveSubmitAction
                        {
                            Title = localizer.GetString("EditButtonText"),
                            Data = new AdaptiveSubmitActionData
                            {
                                AdaptiveActionType = Constants.EditTeamGoalsCommand,
                                TeamId = teamGoalDetailEntity.TeamId,
                                GoalCycleId = goalCycleId,
                            },
                        }),
                        Icon = $"{applicationBasePath}/Artifacts/listIcon.png",
                    });
                }
            }

            if (!isCardForTeamMembers)
            {
                CardAction editTeamGoals = new TaskModuleAction(localizer.GetString("EditButtonText"), new AdaptiveSubmitAction
                {
                    Data = new AdaptiveSubmitActionData
                    {
                        AdaptiveActionType = Constants.EditTeamGoalsCommand,
                        TeamId = teamId,
                        GoalCycleId = goalCycleId,
                    },
                });
                teamGoalDetailsListCard.Buttons.Add(editTeamGoals);
            }

            CardAction alignGoals = new TaskModuleAction(localizer.GetString("AlignGoalsButtonText"), new AdaptiveSubmitAction
            {
                Data = new AdaptiveSubmitActionData
                {
                    AdaptiveActionType = Constants.AlignGoalCommand,
                    TeamId = teamId,
                    GoalCycleId = goalCycleId,
                },
            });
            teamGoalDetailsListCard.Buttons.Add(alignGoals);

            var teamGoalsListCard = new Attachment()
            {
                ContentType = "application/vnd.microsoft.teams.card.list",
                Content = teamGoalDetailsListCard,
            };

            return teamGoalsListCard;
        }

        /// <summary>
        /// Method to show personal goals details card in personal scope.
        /// </summary>
        /// <param name="applicationBasePath">Application base URL of the application.</param>
        /// <param name="personalGoalDetails">Personal goal values entered by user.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="goalCycleId">Goal cycle id to identify that goal cycle is ended or not.</param>
        /// <returns>Returns an attachment of personal goal details card in personal scope.</returns>
        public static Attachment GetPersonalGoalDetailsListCard(string applicationBasePath, IEnumerable<PersonalGoalDetail> personalGoalDetails, IStringLocalizer<Strings> localizer, string goalCycleId)
        {
            personalGoalDetails = personalGoalDetails ?? throw new ArgumentNullException(nameof(personalGoalDetails));
            var goalCycleStartDate = personalGoalDetails.Select(goal => goal.StartDate).First();
            var goalCycleEndDate = personalGoalDetails.Select(goal => goal.EndDate).First();

            // Start date and end date are stored in storage with user specific time zones.
            goalCycleStartDate = DateTime.Parse(goalCycleStartDate, CultureInfo.CurrentCulture)
                .ToString(Constants.ListCardDateTimeFormat, CultureInfo.CurrentCulture);
            goalCycleEndDate = DateTime.Parse(goalCycleEndDate, CultureInfo.CurrentCulture)
                .ToString(Constants.ListCardDateTimeFormat, CultureInfo.CurrentCulture);

            var reminder = personalGoalDetails.Select(goal => goal.ReminderFrequency).First();
            var reminderFrequency = (ReminderFrequency)reminder;
            ListCard personalGoalDetailsListCard = new ListCard
            {
                Title = localizer.GetString("PersonalGoalListCardTitle"),
            };

            personalGoalDetailsListCard.Items.Add(new ListItem
            {
                Title = localizer.GetString("PersonalGoalCardCycleText", goalCycleStartDate, goalCycleEndDate),
                Type = "section",
            });

            foreach (var personalGoalDetailEntity in personalGoalDetails)
            {
                personalGoalDetailsListCard.Items.Add(new ListItem
                {
                    Id = personalGoalDetailEntity.PersonalGoalId,
                    Title = personalGoalDetailEntity.GoalName,
                    Subtitle = reminderFrequency.ToString(),
                    Type = "resultItem",
                    Tap = new TaskModuleAction(localizer.GetString("PersonalGoalEditButtonText"), new AdaptiveSubmitAction
                    {
                        Title = localizer.GetString("PersonalGoalEditButtonText"),
                        Data = new AdaptiveSubmitActionData
                        {
                            AdaptiveActionType = Constants.EditPersonalGoalsCommand,
                            GoalCycleId = goalCycleId,
                        },
                    }),
                    Icon = $"{applicationBasePath}/Artifacts/listIcon.png",
                });
            }

            CardAction editButton = new TaskModuleAction(localizer.GetString("PersonalGoalEditButtonText"), new AdaptiveSubmitAction
            {
                Data = new AdaptiveSubmitActionData
                {
                    AdaptiveActionType = Constants.EditPersonalGoalsCommand,
                    GoalCycleId = goalCycleId,
                },
            });
            personalGoalDetailsListCard.Buttons.Add(editButton);

            CardAction addNoteButton = new TaskModuleAction(localizer.GetString("PersonalGoalAddNoteButtonText"), new AdaptiveSubmitAction
            {
                Data = new AdaptiveSubmitActionData
                {
                    AdaptiveActionType = Constants.AddNoteCommand,
                    GoalCycleId = goalCycleId,
                },
            });
            personalGoalDetailsListCard.Buttons.Add(addNoteButton);

            var personalGoalsListCard = new Attachment()
            {
                ContentType = "application/vnd.microsoft.teams.card.list",
                Content = personalGoalDetailsListCard,
            };

            return personalGoalsListCard;
        }

        /// <summary>
        /// Gets the confirmation card for changing goal alignment from one team to another.
        /// </summary>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="teamId">Team id to align personal goal within specified team.</param>
        /// <returns>Get confirmation card attachment for changing goal alignment.</returns>
        public static Attachment GetAlignmentChangeConfirmationCard(IStringLocalizer<Strings> localizer, string teamId)
        {
            AdaptiveCard alignmentChangeConfirmationCard = new AdaptiveCard(Constants.AdaptiveCardVersion)
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveTextBlock
                    {
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Text = localizer.GetString("ConfirmChangeAlignment"),
                        Wrap = true,
                    },
                },
            };

            alignmentChangeConfirmationCard.Actions = new List<AdaptiveAction>
            {
                new AdaptiveSubmitAction
                {
                    Title = localizer.GetString("ChangeAlignmentNoActionText"),
                    Data = new AdaptiveSubmitActionData
                    {
                        MsTeams = new TaskModuleAction(
                        localizer.GetString("ChangeAlignmentNoActionText"),
                        new AdaptiveSubmitActionData
                        {
                            AdaptiveActionType = Constants.AlignGoalCommand,
                            TeamId = string.Empty,
                        }),
                        AdaptiveActionType = Constants.AlignGoalCommand,
                        TeamId = string.Empty,
                    },
                },
                new AdaptiveSubmitAction
                {
                    Title = localizer.GetString("ChangeAlignmentYesActionText"),
                    Data = new AdaptiveSubmitActionData
                    {
                        MsTeams = new TaskModuleAction(
                        localizer.GetString("ChangeAlignmentYesActionText"),
                        new AdaptiveSubmitActionData
                        {
                            AdaptiveActionType = Constants.AlignGoalCommand,
                            TeamId = teamId,
                        }),
                        AdaptiveActionType = Constants.AlignGoalCommand,
                        TeamId = teamId,
                    },
                },
            };

            return new Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = alignmentChangeConfirmationCard,
            };
        }
    }
}
