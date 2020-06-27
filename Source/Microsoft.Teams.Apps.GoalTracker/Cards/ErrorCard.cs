// <copyright file="ErrorCard.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Cards
{
    using System.Collections.Generic;
    using AdaptiveCards;
    using Microsoft.Bot.Schema;
    using Microsoft.Teams.Apps.GoalTracker.Common;

    /// <summary>
    /// This class is to render the adaptive card with error message.
    /// </summary>
    public static class ErrorCard
    {
        /// <summary>
        /// Construct the adaptive card to render error message.
        /// </summary>
        /// <param name="errorMessage">Error message to be displayed in adaptive card.</param>
        /// <returns>Error card attachment.</returns>
        public static Attachment GetErrorCardAttachment(string errorMessage)
        {
            AdaptiveCard errorCard = new AdaptiveCard(Constants.AdaptiveCardVersion)
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveTextBlock
                    {
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Text = errorMessage,
                        Wrap = true,
                    },
                },
            };
            return new Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = errorCard,
            };
        }
    }
}
