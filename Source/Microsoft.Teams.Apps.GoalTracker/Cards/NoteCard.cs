// <copyright file="NoteCard.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Cards
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using AdaptiveCards;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Localization;
    using Microsoft.Teams.Apps.GoalTracker.Common;
    using Microsoft.Teams.Apps.GoalTracker.Helpers;
    using Microsoft.Teams.Apps.GoalTracker.Models;
    using Newtonsoft.Json;

    /// <summary>
    ///  Class handles personal goal note cards.
    /// </summary>
    public static class NoteCard
    {
        /// <summary>
        /// Represents maximum length for source input field.
        /// </summary>
        private const int SourceInputMaximumLength = 100;

        /// <summary>
        /// Represents maximum length for notes input field.
        /// </summary>
        private const int NotesInputMaximumLength = 1000;

        /// <summary>
        /// Represents minimum height of the container in pixel.
        /// </summary>
        private const int NotesCardContainerHeight = 450;

        /// <summary>
        /// Maximum length of goal name to be shown in the personal goal drop down in add note task module.
        /// </summary>
        private const int TruncateThresholdLength = 80;

        /// <summary>
        /// Get the add note card on command received through personal bot.
        /// </summary>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns>Add note card attachment.</returns>
        public static Attachment GetAddNoteCardOnMessage(IStringLocalizer<Strings> localizer)
        {
            AdaptiveCard addNoteCard = new AdaptiveCard(Constants.AdaptiveCardVersion)
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveTextBlock
                    {
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Text = localizer.GetString("AddNoteTitleText"),
                        Wrap = true,
                    },
                },
                Actions = new List<AdaptiveAction>
                {
                    new AdaptiveSubmitAction
                    {
                        Title = localizer.GetString("AddNoteButtonText"),
                        Data = new AdaptiveSubmitActionData
                        {
                            MsTeams = new TaskModuleAction(localizer.GetString("AddNoteButtonText"), new { data = JsonConvert.SerializeObject(new AdaptiveSubmitActionData { AdaptiveActionType = Constants.AddNoteCommand }) }),
                        },
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
        /// Gets add note card on task module through bot command or button click.
        /// </summary>
        /// <param name="personalGoalNoteDetail">Holds personal goal note detail entity data.</param>
        /// <param name="personalGoalDetail">Holds collection of personal goal detail entity data.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="isNoteEmpty">Determines whether note is empty.</param>
        /// /// <param name="isPersonalGoalEmpty">Determines whether personal goal is empty.</param>
        /// <param name="isNoteCountExceedsTen">Determines whether note count exceeds ten.</param>
        /// <returns>Returns an attachment of card.</returns>
        public static Attachment GetAddNoteCardInTaskModule(PersonalGoalNoteDetail personalGoalNoteDetail, IEnumerable<PersonalGoalDetail> personalGoalDetail, IStringLocalizer<Strings> localizer, bool isNoteEmpty = false, bool isPersonalGoalEmpty = false, bool isNoteCountExceedsTen = false)
        {
            personalGoalDetail = personalGoalDetail ?? throw new ArgumentNullException(nameof(personalGoalDetail));

            List<AdaptiveChoice> personalGoalList = new List<AdaptiveChoice>();
            foreach (var personalGoal in personalGoalDetail)
            {
                string truncatedGoalName = personalGoal.GoalName.Length <= TruncateThresholdLength ? personalGoal.GoalName : personalGoal.GoalName.Substring(0, 80) + "...";
                personalGoalList.Add(new AdaptiveChoice { Title = truncatedGoalName, Value = personalGoal.PersonalGoalId, });
            }

            AdaptiveCard addNoteCard = new AdaptiveCard(new AdaptiveSchemaVersion(1, 0));
            {
                var container = new AdaptiveContainer()
                {
                    PixelMinHeight = NotesCardContainerHeight,
                    Items = new List<AdaptiveElement>
                    {
                        new AdaptiveColumnSet
                        {
                            Columns = new List<AdaptiveColumn>
                            {
                                new AdaptiveColumn
                                {
                                    Width = AdaptiveColumnWidth.Stretch,
                                    Items = new List<AdaptiveElement>
                                    {
                                        new AdaptiveTextBlock
                                        {
                                            Size = AdaptiveTextSize.Medium,
                                            Wrap = true,
                                            Text = localizer.GetString("GoalBucketText"),
                                            Spacing = AdaptiveSpacing.None,
                                        },
                                    },
                                },
                                new AdaptiveColumn
                                {
                                    Width = AdaptiveColumnWidth.Auto,
                                    Items = new List<AdaptiveElement>
                                    {
                                        new AdaptiveTextBlock
                                        {
                                            Size = AdaptiveTextSize.Medium,
                                            Wrap = true,
                                            Text = localizer.GetString("AddNoteEmptyGoalError"),
                                            HorizontalAlignment = AdaptiveHorizontalAlignment.Right,
                                            Color = AdaptiveTextColor.Attention,
                                            IsVisible = isPersonalGoalEmpty,
                                        },
                                    },
                                },
                                new AdaptiveColumn
                                {
                                    Width = AdaptiveColumnWidth.Auto,
                                    Items = new List<AdaptiveElement>
                                    {
                                        new AdaptiveTextBlock
                                        {
                                            Size = AdaptiveTextSize.Medium,
                                            HorizontalAlignment = AdaptiveHorizontalAlignment.Right,
                                            Color = AdaptiveTextColor.Attention,
                                            Text = localizer.GetString("AddNoteMaximumNoteError"),
                                            IsVisible = isNoteCountExceedsTen,
                                            Wrap = true,
                                        },
                                    },
                                },
                            },
                        },
                        new AdaptiveChoiceSetInput
                        {
                            Value = personalGoalNoteDetail?.PersonalGoalId,
                            Choices = personalGoalList,
                            IsMultiSelect = false,
                            Id = "personalgoalid",
                            Style = AdaptiveChoiceInputStyle.Compact,
                        },
                        new AdaptiveColumnSet
                        {
                            Columns = new List<AdaptiveColumn>
                            {
                                new AdaptiveColumn
                                {
                                    Width = AdaptiveColumnWidth.Stretch,
                                    Items = new List<AdaptiveElement>
                                    {
                                        new AdaptiveTextBlock
                                        {
                                            Size = AdaptiveTextSize.Medium,
                                            Text = localizer.GetString("NoteText"),
                                            Spacing = AdaptiveSpacing.None,
                                        },
                                    },
                                },
                                new AdaptiveColumn
                                {
                                    Width = AdaptiveColumnWidth.Auto,
                                    Items = new List<AdaptiveElement>
                                    {
                                        new AdaptiveTextBlock
                                        {
                                            Size = AdaptiveTextSize.Medium,
                                            Wrap = true,
                                            Text = localizer.GetString("AddNoteEmptyNoteError"),
                                            HorizontalAlignment = AdaptiveHorizontalAlignment.Right,
                                            Color = AdaptiveTextColor.Attention,
                                            IsVisible = isNoteEmpty,
                                        },
                                    },
                                },
                            },
                        },
                        new AdaptiveTextInput
                        {
                            Spacing = AdaptiveSpacing.Small,
                            Id = "personalgoalnotedescription",
                            MaxLength = NotesInputMaximumLength,
                            IsMultiline = true,
                            Placeholder = localizer.GetString("AddNotePlaceHolder"),
                            Value = personalGoalNoteDetail?.PersonalGoalNoteDescription,
                        },
                        new AdaptiveTextBlock
                        {
                            Text = localizer.GetString("SourceText"),
                            HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                            Wrap = true,
                        },
                        new AdaptiveTextInput
                        {
                            Spacing = AdaptiveSpacing.Small,
                            Id = "sourcename",
                            MaxLength = SourceInputMaximumLength,
                            Placeholder = localizer.GetString("SourceNamePlaceHolder"),
                            Value = personalGoalNoteDetail?.SourceName,
                        },
                    },
                };
                addNoteCard.Body.Add(container);

                addNoteCard.Actions = new List<AdaptiveAction>
                {
                    new AdaptiveSubmitAction
                    {
                        Title = localizer.GetString("SubmitActionText"),
                        Data = new AdaptiveSubmitActionData
                        {
                            MsTeams = new CardAction
                            {
                                Type = Constants.TaskModuleSubmitType,
                            },
                            AdaptiveActionType = Constants.AddNoteCommand,
                            GoalNoteId = personalGoalNoteDetail.PersonalGoalNoteId,
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

        /// <summary>
        /// Gets the add note submit card.
        /// </summary>
        /// <param name="personalGoalNoteDetail">Instance containing personal goal note related details.</param>
        /// <param name="personalGoalDetail">Instance containing personal goal related details.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns>Attachment having add note submit card.</returns>
        public static Attachment GetsAddNoteSubmitCard(PersonalGoalNoteDetail personalGoalNoteDetail, PersonalGoalDetail personalGoalDetail, IStringLocalizer<Strings> localizer)
        {
            personalGoalDetail = personalGoalDetail ?? throw new ArgumentNullException(nameof(personalGoalDetail));
            personalGoalNoteDetail = personalGoalNoteDetail ?? throw new ArgumentNullException(nameof(personalGoalNoteDetail));
            var isSourceEmpty = string.IsNullOrEmpty(personalGoalNoteDetail?.SourceName) ? true : false;

            AdaptiveCard addCardSubmitCard = new AdaptiveCard(Constants.AdaptiveCardVersion)
            {
                Body = new List<AdaptiveElement>()
                {
                    new AdaptiveTextBlock
                    {
                        Text = localizer.GetString("AddNoteHeading"),
                        Weight = AdaptiveTextWeight.Bolder,
                        Size = AdaptiveTextSize.Large,
                        Wrap = true,
                    },
                    new AdaptiveColumnSet
                    {
                        Spacing = AdaptiveSpacing.Padding,
                        Columns = new List<AdaptiveColumn>
                        {
                            new AdaptiveColumn
                            {
                                Width = "1",
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Size = AdaptiveTextSize.Medium,
                                        Wrap = true,
                                        Text = localizer.GetString("AddNoteGoalNameSubheading"),
                                    },
                                },
                            },
                            new AdaptiveColumn
                            {
                                Width = "4",
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Size = AdaptiveTextSize.Medium,
                                        Wrap = true,
                                        Text = personalGoalDetail.GoalName,
                                        Weight = AdaptiveTextWeight.Bolder,
                                    },
                                },
                            },
                        },
                    },
                    new AdaptiveColumnSet
                    {
                        Spacing = AdaptiveSpacing.Padding,
                        Columns = new List<AdaptiveColumn>
                        {
                            new AdaptiveColumn
                            {
                                Width = "1",
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Size = AdaptiveTextSize.Medium,
                                        Wrap = true,
                                        Text = localizer.GetString("AddNoteNoteSubheading"),
                                        Weight = AdaptiveTextWeight.Bolder,
                                    },
                                },
                            },
                            new AdaptiveColumn
                            {
                                Width = "4",
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Size = AdaptiveTextSize.Medium,
                                        Wrap = true,
                                        Text = personalGoalNoteDetail.PersonalGoalNoteDescription,
                                    },
                                },
                            },
                        },
                    },
                },
            };

            if (!isSourceEmpty)
            {
                addCardSubmitCard.Body.Add(new AdaptiveColumnSet()
                {
                    Spacing = AdaptiveSpacing.None,
                    Columns = new List<AdaptiveColumn>
                    {
                        new AdaptiveColumn
                        {
                            Width = "1",
                            Items = new List<AdaptiveElement>
                            {
                                new AdaptiveTextBlock
                                {
                                    Size = AdaptiveTextSize.Medium,
                                    Wrap = true,
                                    Text = localizer.GetString("AddNoteSourceSubheading"),
                                    Weight = AdaptiveTextWeight.Bolder,
                                },
                            },
                        },
                        new AdaptiveColumn
                        {
                            Width = "4",
                            Items = new List<AdaptiveElement>
                            {
                                new AdaptiveTextBlock
                                {
                                    Size = AdaptiveTextSize.Medium,
                                    Wrap = true,
                                    Text = personalGoalNoteDetail.SourceName,
                                },
                            },
                        },
                    },
                });
            }

            addCardSubmitCard.Body.Add(new AdaptiveColumnSet()
            {
                Spacing = AdaptiveSpacing.None,
                Columns = new List<AdaptiveColumn>
                {
                    new AdaptiveColumn
                    {
                        Width = "1",
                        Items = new List<AdaptiveElement>
                        {
                            new AdaptiveTextBlock
                            {
                                Size = AdaptiveTextSize.Medium,
                                Wrap = true,
                                Text = localizer.GetString("AddNoteDateSubheading"),
                                Weight = AdaptiveTextWeight.Bolder,
                            },
                        },
                    },
                    new AdaptiveColumn
                    {
                        Width = "4",
                        Items = new List<AdaptiveElement>
                        {
                            new AdaptiveTextBlock
                            {
                                Size = AdaptiveTextSize.Medium,
                                Wrap = true,
                                Text = CardHelper.FormatDateStringToAdaptiveCardDateFormat(DateTime.Now.ToString(CultureInfo.CurrentCulture)),
                            },
                        },
                    },
                },
            });

            addCardSubmitCard.Actions = new List<AdaptiveAction>
            {
                new AdaptiveSubmitAction
                {
                    Title = localizer.GetString("AddNoteEditButtonText"),
                    Data = new AdaptiveSubmitActionData
                    {
                        MsTeams = new TaskModuleAction(
                        localizer.GetString("AddNoteEditButtonText"),
                        new
                        {
                            data = JsonConvert.SerializeObject(new AdaptiveSubmitActionData
                            {
                                AdaptiveActionType = Constants.EditNoteCommand,
                                PersonalGoalNoteId = personalGoalNoteDetail.PersonalGoalNoteId,
                                PersonalGoalId = personalGoalNoteDetail.PersonalGoalId,
                                GoalNoteId = personalGoalNoteDetail.AdaptiveCardActivityId,
                            }),
                        }),
                    },
                },
            };

            return new Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = addCardSubmitCard,
            };
        }
    }
}
