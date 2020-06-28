// <copyright file="PersonalGoalController.cs" company="Microsoft">
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
    using Microsoft.Teams.Apps.GoalTracker.Common;
    using Microsoft.Teams.Apps.GoalTracker.Helpers;
    using Microsoft.Teams.Apps.GoalTracker.Models;

    /// <summary>
    /// Controller to handle personal goal API operations.
    /// </summary>
    [Route("api/personalgoals")]
    [ApiController]
    [Authorize]
    public class PersonalGoalController : BaseGoalTrackerController
    {
        /// <summary>
        /// Sends logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<PersonalGoalController> logger;

        /// <summary>
        /// Storage provider for working with personal goal data in storage.
        /// </summary>
        private readonly IPersonalGoalStorageProvider personalGoalStorageProvider;

        /// <summary>
        /// Instance of class that handles card create/update helper methods.
        /// </summary>
        private readonly CardHelper cardHelper;

        /// <summary>
        /// Wrapper class with properties and methods to manage Tasks. Used to run a background task.
        /// </summary>
        private readonly BackgroundTaskWrapper backgroundTaskWrapper;

        /// <summary>
        /// Initializes a new instance of the <see cref="PersonalGoalController"/> class.
        /// </summary>
        /// <param name="confidentialClientApp">Instance of ConfidentialClientApplication class.</param>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="personalGoalStorageProvider">Storage provider for working with personal goal data in storage.</param>
        /// <param name="cardHelper">Instance of class that handles card create/update helper methods.</param>
        /// <param name="backgroundTaskWrapper">Instance of backgroundTaskWrapper to run a background task.</param>
        /// <param name="azureAdOptions">Instance of IOptions to read data from application configuration.</param>
        /// <param name="tokenAcquisitionHelper">Instance of token acquisition helper to access token.</param>
        public PersonalGoalController(
            IConfidentialClientApplication confidentialClientApp,
            ILogger<PersonalGoalController> logger,
            IOptions<AzureAdOptions> azureAdOptions,
            IPersonalGoalStorageProvider personalGoalStorageProvider,
            TokenAcquisitionHelper tokenAcquisitionHelper,
            CardHelper cardHelper,
            BackgroundTaskWrapper backgroundTaskWrapper)
            : base(confidentialClientApp, azureAdOptions, logger, tokenAcquisitionHelper)
        {
            this.logger = logger;
            this.personalGoalStorageProvider = personalGoalStorageProvider;
            this.cardHelper = cardHelper;
            this.backgroundTaskWrapper = backgroundTaskWrapper;
        }

        /// <summary>
        /// Get all personal goal details of a user by user Azure Active Directory object id.
        /// </summary>
        /// <returns>Returns personal goal details as obtained from storage.</returns>
        [HttpGet]
        public async Task<IActionResult> GetPersonalGoalDetailsAsync()
        {
            try
            {
                this.logger.LogInformation("Initiated call for fetching personal goal details from storage");
                var personalGoalDetails = await this.personalGoalStorageProvider.GetPersonalGoalDetailsByUserAadObjectIdAsync(this.UserObjectId);
                this.logger.LogInformation("GET call for fetching personal goal details from storage is successful");
                return this.Ok(personalGoalDetails);
            }
#pragma warning disable CA1031 // Catching all generic exceptions in order to log exception details in logger
            catch (Exception ex)
#pragma warning restore CA1031 // Catching all generic exceptions in order to log exception details in logger
            {
                this.logger.LogError(ex, "Error while getting personal goal details");
                throw;
            }
        }

        /// <summary>
        /// Get details of a personal goal by personal goal Id.
        /// </summary>
        /// <param name="personalGoalId">Unique identifier of personal goal which user wants to edit.</param>
        /// <returns>Returns personal goal details received from storage.</returns>
        [HttpGet("{personalGoalId}")]
        public async Task<IActionResult> GetPersonalGoalDetailByGoalIdAsync(string personalGoalId)
        {
            try
            {
                if (!Guid.TryParse(personalGoalId, out var validPersonalGoalId))
                {
                    this.logger.LogError(StatusCodes.Status400BadRequest, $"Personal goal id:{personalGoalId} is not a valid GUID.");
                    return this.BadRequest($"Personal goal id:{personalGoalId} is not a valid GUID.");
                }

                this.logger.LogInformation("Initiated call for fetching personal goal details from storage");
                var personalGoalDetail = await this.personalGoalStorageProvider.GetPersonalGoalDetailByGoalIdAsync(personalGoalId, this.UserObjectId);
                this.logger.LogInformation("GET call for fetching personal goal details from storage is successful");
                return this.Ok(personalGoalDetail);
            }
#pragma warning disable CA1031 // Catching all generic exceptions in order to log exception details in logger
            catch (Exception ex)
#pragma warning restore CA1031 // Catching all generic exceptions in order to log exception details in logger
            {
                this.logger.LogError(ex, "Error while getting personal goal details");
                throw;
            }
        }

        /// <summary>
        /// Put call to update a personal goal detail in storage.
        /// </summary>
        /// <param name="personalGoalData">Class contains detail of personal goal to be saved or updated.</param>
        /// <returns>Returns true for successful operation.</returns>
        [HttpPatch("{personalGoalId}")]
        public async Task<IActionResult> UpdatePersonalGoalDetailAsync(PersonalGoalDetail personalGoalData)
        {
            try
            {
#pragma warning disable CA1062 // Patch details are validated by model validations for null check and is responded with bad request status
                var validationErrorResponse = this.ValidatePersonalGoal(personalGoalData);
#pragma warning restore CA1062 // Patch details are validated by model validations for null check and is responded with bad request status
                if (validationErrorResponse.StatusCode != StatusCodes.Status200OK)
                {
                    return validationErrorResponse;
                }

                var existingGoalDetail = await this.personalGoalStorageProvider.GetPersonalGoalDetailByGoalIdAsync(personalGoalData.PersonalGoalId, this.UserObjectId);

                if (existingGoalDetail == null)
                {
                    this.logger.LogError(StatusCodes.Status404NotFound, $"The personal goal user trying to update does not exist. Personal goal id: {personalGoalData.PersonalGoalId} ");
                    return this.NotFound("The personal goal user trying to update does not exist.");
                }

                this.logger.LogInformation("Initiated call to personal goal storage provider.");
                var result = await this.personalGoalStorageProvider.CreateOrUpdatePersonalGoalDetailAsync(personalGoalData);
                this.logger.LogInformation("PATCH call for saving personal goal detail in storage is successful");

                if (!result)
                {
                    this.logger.LogError($"Could not save or update goal data received with personal goal id: {personalGoalData.PersonalGoalId}.");
                    return this.StatusCode(StatusCodes.Status500InternalServerError, "Could not save or update goal data received.");
                }

                // Update list card of personal bot. Enqueue task to task wrapper and it will be executed by goal background service.
                this.backgroundTaskWrapper.Enqueue(this.cardHelper.UpdatePersonalGoalListCardAsync(personalGoalData));
                return this.Ok(result);
            }
#pragma warning disable CA1031 // Catching all generic exceptions in order to log exception details in logger
            catch (Exception ex)
#pragma warning restore CA1031 // Catching all generic exceptions in order to log exception details in logger
            {
                this.logger.LogError(ex, "Error while saving personal goal detail");
                throw;
            }
        }

        /// <summary>
        /// Post call to save or update personal goal details in storage.
        /// </summary>
        /// <param name="personalGoalsData">Class contains details of personal goal to be saved or updated.</param>
        /// <returns>Returns true for successful operation.</returns>
        [HttpPost]
        public async Task<IActionResult> SaveOrUpdatePersonalGoalDetailsAsync(IEnumerable<PersonalGoalDetail> personalGoalsData)
        {
            try
            {
#pragma warning disable CA1062 // Post details are validated by model validations for null check and is responded with bad request status
                var validationResponse = this.ValidatePersonalGoals(personalGoalsData);
#pragma warning restore CA1062 // Post details are validated by model validations for null check and is responded with bad request status
                if (validationResponse.StatusCode != StatusCodes.Status200OK)
                {
                    return validationResponse;
                }

                this.logger.LogInformation("Initiated call to personal goal storage provider.");
                var result = await this.personalGoalStorageProvider.CreateOrUpdatePersonalGoalDetailsAsync(personalGoalsData);
                this.logger.LogInformation("POST call for saving personal goal details in storage is successful");

                if (!result)
                {
                    this.logger.LogError($"Could not save or update goals data received for goal cycle: {personalGoalsData.First().GoalCycleId}.");
                    return this.StatusCode(StatusCodes.Status500InternalServerError, "Could not save or update goals data received.");
                }

                return this.Ok(result);
            }
#pragma warning disable CA1031 // Catching all generic exceptions in order to log exception details in logger
            catch (Exception ex)
#pragma warning restore CA1031 // Catching all generic exceptions in order to log exception details in logger
            {
                this.logger.LogError(ex, "Error while saving personal goal details");
                throw;
            }
        }

        /// <summary>
        /// Call to delete specified personal goal detail from storage.
        /// </summary>
        /// <param name="personalGoalId">Unique identifier of personal goal to be deleted.</param>
        /// <returns>Returns true for successful operation of deleting personal goal.</returns>
        [HttpDelete("{personalGoalId}")]
        public async Task<IActionResult> DeletePersonalGoalDetailAsync(string personalGoalId)
        {
            try
            {
                if (!Guid.TryParse(personalGoalId, out var validPersonalGoalId))
                {
                    this.logger.LogError(StatusCodes.Status400BadRequest, $"Personal goal id:{personalGoalId} is not a valid GUID.");
                    return this.BadRequest($"Personal goal id:{personalGoalId} is not a valid GUID.");
                }

                this.logger.LogInformation("Initiated call to personal goal storage provider service to delete personal goal detail.");
                var existingGoalDetail = await this.personalGoalStorageProvider.GetPersonalGoalDetailByGoalIdAsync(personalGoalId, this.UserObjectId);
                if (existingGoalDetail == null)
                {
                    this.logger.LogError(StatusCodes.Status404NotFound, $"The personal goal with personal goal id {personalGoalId} user trying to delete does not exist.");
                    return this.NotFound("The personal goal user trying to delete does not exist.");
                }

                // Update IsDeleted flag in storage to false. Background service will delete all data weekly from storage where IsDeleted = true
                // This is required so that search index will show correct aligned goal count as it takes time to refresh documents in index.
                existingGoalDetail.IsDeleted = true;
                existingGoalDetail.IsActive = false;
                var result = await this.personalGoalStorageProvider.CreateOrUpdatePersonalGoalDetailAsync(existingGoalDetail);
                this.logger.LogInformation("DELETE call for deleting personal goal detail in storage is successful");

                if (!result)
                {
                    this.logger.LogError(StatusCodes.Status500InternalServerError, $"Could not delete goal data received with personal goal id {personalGoalId}.");
                    return this.StatusCode(StatusCodes.Status500InternalServerError, "Could not delete goal data received.");
                }

                // Update list card of personal bot. Enqueue task to task wrapper and it will be executed by goal background service.
                this.backgroundTaskWrapper.Enqueue(this.cardHelper.UpdatePersonalGoalListCardAsync(existingGoalDetail));
                return this.Ok(result);
            }
#pragma warning disable CA1031 // Catching all generic exceptions in order to log exception details in logger
            catch (Exception ex)
#pragma warning restore CA1031 // Catching all generic exceptions in order to log exception details in logger
            {
                this.logger.LogError(ex, "Error while deleting personal goal detail.");
                throw;
            }
        }

        /// <summary>
        /// Validates a personal goal data received from client application is valid.
        /// </summary>
        /// <param name="personalGoalData">Details of a personal goal to be saved/updated in storage</param>
        /// <returns> Returns response representing whether data received is valid or not</returns>
        private ObjectResult ValidatePersonalGoal(PersonalGoalDetail personalGoalData)
        {
            if (!Guid.TryParse(personalGoalData.PersonalGoalId, out _))
            {
                this.logger.LogError(StatusCodes.Status400BadRequest, $"Personal goal id:{personalGoalData.PersonalGoalId} is not a valid GUID.");
                return this.BadRequest($"Personal goal id:{personalGoalData.PersonalGoalId} is not a valid GUID.");
            }
            else if (personalGoalData.UserAadObjectId != this.UserObjectId)
            {
                this.logger.LogError(StatusCodes.Status403Forbidden, $"User {personalGoalData.UserAadObjectId} is forbidden to perform this operation.");
                return this.StatusCode(StatusCodes.Status403Forbidden, $"User {personalGoalData.UserAadObjectId} is forbidden to perform this operation.");
            }
            else if (personalGoalData.CreatedBy != this.HttpContext.User.Identity.Name)
            {
                this.logger.LogError(StatusCodes.Status403Forbidden, $"User {personalGoalData.CreatedBy} is forbidden to perform this operation.");
                return this.StatusCode(StatusCodes.Status403Forbidden, $"User {personalGoalData.CreatedBy} is forbidden to perform this operation.");
            }
            else if (DateTime.Parse(personalGoalData.StartDate, CultureInfo.InvariantCulture) >= DateTime.Parse(personalGoalData.EndDate, CultureInfo.InvariantCulture))
            {
                this.logger.LogError(StatusCodes.Status400BadRequest, $"Personal goal start date is greater than end date.");
                return this.BadRequest("Personal goal start date is greater than end date.");
            }
            else if (DateTime.Parse(personalGoalData.EndDate, CultureInfo.InvariantCulture).ToUniversalTime() < DateTime.Now.ToUniversalTime())
            {
                this.logger.LogError(StatusCodes.Status400BadRequest, $"Personal goal end date is smaller than current date.");
                return this.BadRequest("Personal goal end date is smaller than current date.");
            }

            this.logger.LogInformation(StatusCodes.Status200OK, $"Personal goal detail received is valid.");
            return this.Ok("Personal goal detail received is valid.");
        }

        /// <summary>
        /// Validates personal goal collection received from client application is valid.
        /// </summary>
        /// <param name="personalGoalDetails">Collection of personal goal details to be saved/updated in storage</param>
        /// <returns> Returns response representing whether data received is valid or not</returns>
        private ObjectResult ValidatePersonalGoals(IEnumerable<PersonalGoalDetail> personalGoalDetails)
        {
            if (personalGoalDetails.Count() > Constants.MaximumNumberOfGoals)
            {
                this.logger.LogError(StatusCodes.Status400BadRequest, $"Cannot add personal goals more than {Constants.MaximumNumberOfGoals}.");
                return this.BadRequest($"Cannot add personal goals more than {Constants.MaximumNumberOfGoals}.");
            }

            foreach (var personalGoalDetail in personalGoalDetails)
            {
                var response = this.ValidatePersonalGoal(personalGoalDetail);
                if (response.StatusCode != StatusCodes.Status200OK)
                {
                    return response;
                }
            }

            this.logger.LogInformation(StatusCodes.Status200OK, $"Personal goals collection received is valid.");
            return this.Ok("Personal goals collection received is valid.");
        }
    }
}