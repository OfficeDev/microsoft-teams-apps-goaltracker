// <copyright file="TeamGoalController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Controllers
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Http;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Identity.Client;
    using Microsoft.Teams.Apps.GoalTracker.Authentication;
    using Microsoft.Teams.Apps.GoalTracker.Common;
    using Microsoft.Teams.Apps.GoalTracker.Helpers;
    using Microsoft.Teams.Apps.GoalTracker.Models;

    /// <summary>
    /// Controller to handle team goal API operations.
    /// </summary>
    [Route("api/teamgoals")]
    [ApiController]
    [Authorize]
    public class TeamGoalController : BaseGoalTrackerController
    {
        /// <summary>
        /// Sends logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<TeamGoalController> logger;

        /// <summary>
        /// Storage provider for working with team goal data in storage.
        /// </summary>
        private readonly ITeamGoalStorageProvider teamGoalStorageProvider;

        /// <summary>
        /// Instance of graphUtilityHelper to access Microsoft Graph API.
        /// </summary>
        private GraphUtilityHelper graphUtilityHelper;

        /// <summary>
        /// Initializes a new instance of the <see cref="TeamGoalController"/> class.
        /// </summary>
        /// <param name="confidentialClientApp">Instance of ConfidentialClientApplication class.</param>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="azureAdOptions">Instance of IOptions to read data from application configuration.</param>
        /// <param name="teamGoalStorageProvider">Storage provider for working with team goal data in Microsoft Azure Table storage</param>
        /// <param name="tokenAcquisitionHelper">Instance of token acquisition helper to access token.</param>
        public TeamGoalController(
            IConfidentialClientApplication confidentialClientApp,
            ILogger<TeamGoalController> logger,
            IOptions<AzureAdOptions> azureAdOptions,
            ITeamGoalStorageProvider teamGoalStorageProvider,
            TokenAcquisitionHelper tokenAcquisitionHelper)
            : base(confidentialClientApp, azureAdOptions, logger, tokenAcquisitionHelper)
        {
            this.logger = logger;
            this.teamGoalStorageProvider = teamGoalStorageProvider;
        }

        /// <summary>
        /// Get team goal details by Microsoft Teams' team Id.
        /// </summary>
        /// <param name="teamId">Team id for which team goal details need to be fetched.</param>
        /// <returns>Returns team goal details.</returns>
        [HttpGet]
        [Authorize(PolicyNames.MustBePartOfTeamPolicy)]
        public async Task<IActionResult> GetTeamGoalDetailsByTeamIdAsync(string teamId)
        {
            try
            {
                this.logger.LogInformation("Initiated call for fetching team goal details from storage.");
                var teamGoalDetails = await this.teamGoalStorageProvider.GetTeamGoalDetailsByTeamIdAsync(teamId);
                this.logger.LogInformation("GET call for fetching team goal details from storage is successful.");
                return this.Ok(teamGoalDetails);
            }
#pragma warning disable CA1031 // Catching all generic exceptions in order to log exception details in logger
            catch (Exception ex)
#pragma warning restore CA1031 // Catching all generic exceptions in order to log exception details in logger
            {
                this.logger.LogError(ex, "Error while getting team goal details.");
                throw;
            }
        }

        /// <summary>
        /// Method to check if logged in user is team owner.
        /// </summary>
        /// <param name="teamGroupId">AAD group id of the team in which bot is installed.</param>
        /// <returns>Returns boolean value that represents whether logged in user is team owner or not.</returns>
        [HttpGet("{teamGroupId}/checkteamowner")]
        public async Task<IActionResult> CheckUserIsTeamOwnerAsync(string teamGroupId)
        {
            try
            {
                try
                {
                    var teamOwnerResponse = await this.ValidateIfUserIsTeamOwnerAsync(teamGroupId);
                    return teamOwnerResponse;
                }
                catch (Exception ex)
                {
                    this.logger.LogError(ex, "Error while validating if user is a team owner.");
                    throw;
                }
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while getting team owner details from graph API.");
                throw;
            }
        }

        /// <summary>
        /// Get specific team goal detail by unique team goal id.
        /// </summary>
        /// <param name="teamGoalId">Team goal id for which team goal details need to be fetched.</param>
        /// <param name="teamId">Team id for which team goal details need to be fetched.</param>
        /// <returns>Returns specific team goal detail.</returns>
        [HttpGet("goal")]
        [Authorize(PolicyNames.MustBePartOfTeamPolicy)]
        public async Task<IActionResult> GetTeamGoalDetailByTeamGoalIdAsync(string teamGoalId, string teamId)
        {
            try
            {
                if (!Guid.TryParse(teamGoalId, out var validTeamGoalId))
                {
                    this.logger.LogError(StatusCodes.Status400BadRequest, $"Team goal id:{teamGoalId} is not a valid GUID.");
                    return this.BadRequest($"Team goal id:{teamGoalId} is not a valid GUID.");
                }

                this.logger.LogInformation("Initiated call for fetching team goal detail from storage.");
                var teamGoalDetail = await this.teamGoalStorageProvider.GetTeamGoalDetailByTeamGoalIdAsync(teamGoalId, teamId);
                this.logger.LogInformation("GET call for fetching team goal detail from storage is successful.");
                return this.Ok(teamGoalDetail);
            }
#pragma warning disable CA1031 // Catching all generic exceptions in order to log exception details in logger
            catch (Exception ex)
#pragma warning restore CA1031 // Catching all generic exceptions in order to log exception details in logger
            {
                this.logger.LogError(ex, $"Error while getting team goal detail from storage for team goal id {teamGoalId}.");
                throw;
            }
        }

        /// <summary>
        /// Post call to save or update team goal details in storage.
        /// </summary>
        /// <param name="teamGroupId">AAD group id of the team in which bot is installed.</param>
        /// <param name="teamGoalsData">Class contains details of team goal to be saved or updated.</param>
        /// <returns>Returns true for successful operation.</returns>
        [HttpPost("{teamGroupId}")]
        public async Task<IActionResult> SaveOrUpdateTeamGoalDetailsAsync(string teamGroupId, [FromBody] IEnumerable<TeamGoalDetail> teamGoalsData)
        {
            try
            {
                var teamOwnerResponse = await this.ValidateIfUserIsTeamOwnerAsync(teamGroupId);
                if (teamOwnerResponse.StatusCode != StatusCodes.Status200OK)
                {
                    return teamOwnerResponse;
                }

#pragma warning disable CA1062 // Post details are validated by model validations for null check and is responded with bad request status
                var validationResponse = this.ValidateTeamGoals(teamGoalsData);
#pragma warning restore CA1062 // Post details are validated by model validations for null check and is responded with bad request status
                if (validationResponse.StatusCode != StatusCodes.Status200OK)
                {
                    return validationResponse;
                }

                this.logger.LogInformation("Initiated call to team goal storage provider service to save team goal details.");
                var result = await this.teamGoalStorageProvider.CreateOrUpdateTeamGoalDetailsAsync(teamGoalsData);
                if (!result)
                {
                    this.logger.LogError($"Could not save or update goals data received.");
                    return this.StatusCode(StatusCodes.Status500InternalServerError, "Could not save or update goals data received.");
                }

                this.logger.LogInformation("POST call for saving team goal details in storage is successful.");
                return this.Ok(result);
            }
#pragma warning disable CA1031 // Catching all generic exceptions in order to log exception details in logger
            catch (Exception ex)
#pragma warning restore CA1031 // Catching all generic exceptions in order to log exception details in logger
            {
                this.logger.LogError(ex, "Error while saving team goal details.");
                throw;
            }
        }

        /// <summary>
        /// Validates team goal collection received from client application is valid.
        /// </summary>
        /// <param name="teamGoalDetails">Collection of team goal details to be saved/updated/deleted in storage</param>
        /// <returns> Returns response representing whether data received is valid or not</returns>
        private ObjectResult ValidateTeamGoals(IEnumerable<TeamGoalDetail> teamGoalDetails)
        {
            if (teamGoalDetails.Count() > Constants.MaximumNumberOfGoals)
            {
                this.logger.LogError(StatusCodes.Status400BadRequest, $"Cannot add personal goals more than {Constants.MaximumNumberOfGoals}.");
                return this.BadRequest($"Cannot add personal goals more than {Constants.MaximumNumberOfGoals}.");
            }

            foreach (var teamGoalDetail in teamGoalDetails)
            {
                if (!Guid.TryParse(teamGoalDetail.TeamGoalId, out _))
                {
                    this.logger.LogError(StatusCodes.Status400BadRequest, $"Team goal id:{teamGoalDetail.TeamGoalId} is not a valid GUID.");
                    return this.BadRequest($"Team goal id:{teamGoalDetail.TeamGoalId} is not a valid GUID.");
                }
                else if (DateTime.Parse(teamGoalDetail.TeamGoalStartDate, CultureInfo.InvariantCulture) >= DateTime.Parse(teamGoalDetail.TeamGoalEndDate, CultureInfo.InvariantCulture))
                {
                    this.logger.LogError(StatusCodes.Status400BadRequest, "Team goal start date is greater than end date.");
                    return this.BadRequest("Team goal start date is greater than end date.");
                }
                else if (DateTime.Parse(teamGoalDetail.TeamGoalEndDate, CultureInfo.InvariantCulture).ToUniversalTime() < DateTime.Now.ToUniversalTime())
                {
                    this.logger.LogError(StatusCodes.Status400BadRequest, "Team goal end date is smaller than current date.");
                    return this.BadRequest("Team goal end date is smaller than current date.");
                }
            }

            this.logger.LogInformation("Team goals collection received is valid.");
            return this.Ok("Team goals collection received is valid.");
        }

        /// <summary>
        /// Validate if user is a team owner using Microsoft Graph API.
        /// </summary>
        /// <param name="teamGroupId">AAD group id of the team in which bot is installed.</param>
        /// <returns>Returns response that represents whether logged in user is team owner or not.</returns>
        private async Task<ObjectResult> ValidateIfUserIsTeamOwnerAsync(string teamGroupId)
        {
            try
            {
                string accessToken = await this.GetAccessTokenAsync();
                if (string.IsNullOrEmpty(accessToken))
                {
                    this.logger.LogError("Token to access graph API is null.");
                    return this.BadRequest("Token to access graph API is null.");
                }

                this.graphUtilityHelper = new GraphUtilityHelper(accessToken);
                var teamOwnerDetails = await this.graphUtilityHelper.GetTeamOwnerDetailsAsync(teamGroupId);
                if (teamOwnerDetails == null)
                {
                    this.logger.LogError(StatusCodes.Status404NotFound, "No data received corresponding to team owner.");
                    return this.NotFound("No data received corresponding to team owner.");
                }

                var isTeamOwner = teamOwnerDetails.ToList().Exists(userId => userId.TeamOwnerId == this.UserObjectId);
                if (!isTeamOwner)
                {
                    this.logger.LogError(StatusCodes.Status403Forbidden, $"User {this.UserObjectId} is not a team owner. User is forbidden to access team goal details.");
                    return this.StatusCode(StatusCodes.Status403Forbidden, $"User {this.UserObjectId} is not a team owner. User is forbidden to access team goal details.");
                }

                return this.Ok(isTeamOwner);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "Error while getting team owner details from graph API.");
                throw;
            }
        }
    }
}
