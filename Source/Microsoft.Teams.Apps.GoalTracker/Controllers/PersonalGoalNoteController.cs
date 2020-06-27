// <copyright file="PersonalGoalNoteController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.GoalTracker.Controllers
{
    using System;
    using System.Collections.Generic;
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
    /// Controller to handle personal goal note API operations.
    /// </summary>
    [Route("api/notes")]
    [ApiController]
    [Authorize]
    public class PersonalGoalNoteController : BaseGoalTrackerController
    {
        /// <summary>
        /// Sends logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<PersonalGoalNoteController> logger;

        /// <summary>
        /// Storage provider for working with personal goal note data in storage.
        /// </summary>
        private readonly IPersonalGoalNoteStorageProvider personalGoalNoteStorageProvider;

        /// <summary>
        /// Instance of class that handles card create/update helper methods.
        /// </summary>
        private readonly CardHelper cardHelper;

        /// <summary>
        /// Wrapper class with properties and methods to manage Tasks. Used to run a background task.
        /// </summary>
        private readonly BackgroundTaskWrapper backgroundTaskWrapper;

        /// <summary>
        /// Initializes a new instance of the <see cref="PersonalGoalNoteController"/> class.
        /// </summary>
        /// <param name="confidentialClientApp">Instance of ConfidentialClientApplication class.</param>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="azureAdOptions">Instance of IOptions to read data from application configuration.</param>
        /// <param name="personalGoalNoteStorageProvider">Storage provider for working with team goal data in Microsoft Azure Table storage</param>
        /// <param name="tokenAcquisitionHelper">Instance of token acquisition helper to access token.</param>
        /// <param name="cardHelper">Instance of class that handles card create/update helper methods.</param>
        /// <param name="backgroundTaskWrapper">Instance of backgroundTaskWrapper to run a background task.</param>
        public PersonalGoalNoteController(
            IConfidentialClientApplication confidentialClientApp,
            ILogger<PersonalGoalNoteController> logger,
            IOptions<AzureAdOptions> azureAdOptions,
            IPersonalGoalNoteStorageProvider personalGoalNoteStorageProvider,
            TokenAcquisitionHelper tokenAcquisitionHelper,
            CardHelper cardHelper,
            BackgroundTaskWrapper backgroundTaskWrapper)
            : base(confidentialClientApp, azureAdOptions, logger, tokenAcquisitionHelper)
        {
            this.logger = logger;
            this.personalGoalNoteStorageProvider = personalGoalNoteStorageProvider;
            this.cardHelper = cardHelper;
            this.backgroundTaskWrapper = backgroundTaskWrapper;
        }

        /// <summary>
        /// Get personal goal note count details by user object id.
        /// </summary>
        /// <returns>Returns personal goal note details.</returns>
        [HttpGet("count")]
        public async Task<IActionResult> GetPersonalGoalNoteCountByUserAadObjectIdAsync()
        {
            try
            {
                this.logger.LogInformation("Initiated call for fetching personal goal note details from storage");

                var personalGoalNoteDetails = await this.personalGoalNoteStorageProvider.GetPersonalGoalNoteDetailsByUserAadObjectIdAsync(this.UserObjectId);
                this.logger.LogInformation("GET call for fetching personal goal note details from storage is successful");

                if (personalGoalNoteDetails != null)
                {
                    return this.Ok(personalGoalNoteDetails
                        .GroupBy(personalGoalNote => personalGoalNote.PersonalGoalId)
                        .Select(personalGoalNote => new
                        {
                            PersonalGoalId = personalGoalNote.Key,
                            NotesCount = personalGoalNote.Count(),
                        }));
                }

                return this.Ok(personalGoalNoteDetails);
            }
#pragma warning disable CA1031 // Catching all generic exceptions in order to log exception details in logger
            catch (Exception ex)
#pragma warning restore CA1031 // Catching all generic exceptions in order to log exception details in logger
            {
                this.logger.LogError(ex, "Error while getting personal goal note details");
                throw;
            }
        }

        /// <summary>
        /// Get personal goal note details added for specific personal goal.
        /// </summary>
        /// <param name="personalGoalId">Unique identifier of personal goal for which notes are fetched.</param>
        /// <returns>Returns personal goal note details.</returns>
        [HttpGet("goal/{personalGoalId}")]
        public async Task<IActionResult> GetPersonalGoalNoteDetailsByPersonalGoalIdAsync(string personalGoalId)
        {
            try
            {
                if (!Guid.TryParse(personalGoalId, out var validPersonalGoalId))
                {
                    this.logger.LogError(StatusCodes.Status400BadRequest, $"Personal goal id:{personalGoalId} is not a valid GUID.");
                    return this.BadRequest($"Personal goal id:{personalGoalId} is not a valid GUID.");
                }

                this.logger.LogInformation("Initiated call for fetching personal goal note details from storage");
                var personalGoalNoteDetails = await this.personalGoalNoteStorageProvider.GetPersonalGoalNoteDetailsAsync(personalGoalId, this.UserObjectId);
                this.logger.LogInformation("GET call for fetching personal goal note details from storage is successful");
                return this.Ok(personalGoalNoteDetails);
            }
#pragma warning disable CA1031 // Catching all generic exceptions in order to log exception details in logger
            catch (Exception ex)
#pragma warning restore CA1031 // Catching all generic exceptions in order to log exception details in logger
            {
                this.logger.LogError(ex, "Error while getting personal goal note details");
                throw;
            }
        }

        /// <summary>
        /// Put call to save or update personal goal note details in storage.
        /// </summary>
        /// <param name="personalGoalNoteDetails">Class contains details of personal goal note to be saved or updated.</param>
        /// <returns>Returns true for successful operation.</returns>
        [HttpPut]
        public async Task<IActionResult> UpdatePersonalGoalNoteDetailsAsync(IEnumerable<PersonalGoalNoteDetail> personalGoalNoteDetails)
        {
            try
            {
#pragma warning disable CA1062 // Put details are validated by model validations for null check and is responded with bad request status
                var validationResponse = this.ValidatePersonalGoalNotes(personalGoalNoteDetails);
#pragma warning restore CA1062 // Put details are validated by model validations for null check and is responded with bad request status
                if (validationResponse.StatusCode != StatusCodes.Status200OK)
                {
                    return validationResponse;
                }

                this.logger.LogInformation("Initiated call to personal goal note storage provider.");
                var result = await this.personalGoalNoteStorageProvider.UpdatePersonalGoalNoteDetailsAsync(personalGoalNoteDetails);

                if (!result)
                {
                    this.logger.LogError(StatusCodes.Status500InternalServerError, $"Could not save or update notes data received.");
                    return this.StatusCode(StatusCodes.Status500InternalServerError, "Could not save or update notes data received.");
                }

                this.logger.LogInformation("PUT call for saving personal goal note details in storage is successful.");

                // Update personal goal note card of personal bot. Enqueue task to task wrapper and it will be executed by goal background service.
                this.backgroundTaskWrapper.Enqueue(this.cardHelper.UpdatePersonalNoteCardAsync(personalGoalNoteDetails));
                return this.Ok(result);
            }
#pragma warning disable CA1031 // Catching all generic exceptions in order to log exception details in logger
            catch (Exception ex)
#pragma warning restore CA1031 // Catching all generic exceptions in order to log exception details in logger
            {
                this.logger.LogError(ex, "Error while updating personal goal note details.");
                throw;
            }
        }

        /// <summary>
        /// Delete call to delete specified personal goal note detail from storage.
        /// </summary>
        /// <param name="personalGoalNotesIds">Class contains details of personal goal note ids to be deleted.</param>
        /// <returns>Returns true for successful operation.</returns>
        [HttpDelete]
        public async Task<IActionResult> DeletePersonalGoalNoteDetailsAsync(IEnumerable<string> personalGoalNotesIds)
        {
            try
            {
                if (personalGoalNotesIds == null || !personalGoalNotesIds.Any())
                {
                    this.logger.LogError(StatusCodes.Status400BadRequest, $"No personal goal note data received to be deleted from storage.");
                    return this.BadRequest("No personal goal note data received to be deleted from storage.");
                }

                // Get data from storage for personal goal note ids received.
                List<PersonalGoalNoteDetail> personalGoalNoteDetails = new List<PersonalGoalNoteDetail>();
                foreach (var personalGoalNoteId in personalGoalNotesIds)
                {
                    if (!Guid.TryParse(personalGoalNoteId, out var validPersonalGoalId))
                    {
                        this.logger.LogError(StatusCodes.Status400BadRequest, $"Personal goal note id:{personalGoalNoteId} is not a valid GUID.");
                        return this.BadRequest($"Personal goal note id:{personalGoalNoteId} is not a valid GUID.");
                    }

                    var personalGoalNoteDetail = await this.personalGoalNoteStorageProvider.GetPersonalGoalNoteDetailAsync(personalGoalNoteId, this.UserObjectId);
                    if (personalGoalNoteDetail == null)
                    {
                        this.logger.LogError(StatusCodes.Status404NotFound, $"Personal goal note with id {personalGoalNoteId} not found in storage.");
                        return this.NotFound($"Personal goal note with id {personalGoalNoteId} not found in storage.");
                    }

                    personalGoalNoteDetails.Add(personalGoalNoteDetail);
                }

                this.logger.LogInformation("Initiated call to personal goal note storage provider service to delete personal goal note details");
                var result = await this.personalGoalNoteStorageProvider.DeletePersonalGoalNoteDetailsAsync(personalGoalNoteDetails);

                if (!result)
                {
                    this.logger.LogError(StatusCodes.Status500InternalServerError, $"Could not delete notes data received.");
                    return this.StatusCode(StatusCodes.Status500InternalServerError, "Could not delete notes data received.");
                }

                this.logger.LogInformation("Delete call for deleting personal goal note details in storage is successful");

                // Update personal goal note card of personal bot. Enqueue task to task wrapper and it will be executed by goal background service.
                this.backgroundTaskWrapper.Enqueue(this.cardHelper.UpdatePersonalNoteCardAsync(personalGoalNoteDetails));
                return this.Ok(result);
            }
#pragma warning disable CA1031 // Catching all generic exceptions in order to log exception details in logger
            catch (Exception ex)
#pragma warning restore CA1031 // Catching all generic exceptions in order to log exception details in logger
            {
                this.logger.LogError(ex, "Error while deleting personal goal note details");
                throw;
            }
        }

        /// <summary>
        /// Validates if personal goal notes collection received from client application is valid.
        /// </summary>
        /// <param name="personalGoalNoteDetails">Collection of personal goal note details to be saved/updated/deleted in storage</param>
        /// <returns> Returns response representing whether data received is valid or not</returns>
        private ObjectResult ValidatePersonalGoalNotes(IEnumerable<PersonalGoalNoteDetail> personalGoalNoteDetails)
        {
            if (personalGoalNoteDetails.Count() > Constants.MaximumNumberOfNotes)
            {
                this.logger.LogError(StatusCodes.Status400BadRequest, $" add personal goal notes more than {Constants.MaximumNumberOfNotes}.");
                return this.BadRequest($"Cannot add personal goal notes more than {Constants.MaximumNumberOfNotes}.");
            }

            foreach (var personalGoalNoteDetail in personalGoalNoteDetails)
            {
                if (!Guid.TryParse(personalGoalNoteDetail.PersonalGoalNoteId, out _))
                {
                    this.logger.LogError(StatusCodes.Status400BadRequest, $"Personal goal note id:{personalGoalNoteDetail.PersonalGoalNoteId} is not a valid GUID.");
                    return this.BadRequest($"Personal goal note id:{personalGoalNoteDetail.PersonalGoalNoteId} is not a valid GUID.");
                }
                else if (personalGoalNoteDetail.UserAadObjectId != this.UserObjectId)
                {
                    this.logger.LogError(StatusCodes.Status403Forbidden, $"User {personalGoalNoteDetail.UserAadObjectId} is forbidden to perform this operation.");
                    return this.StatusCode(StatusCodes.Status403Forbidden, $"User {personalGoalNoteDetail.UserAadObjectId} is forbidden to perform this operation.");
                }
                else if (personalGoalNoteDetail.CreatedBy != this.HttpContext.User.Identity.Name)
                {
                    this.logger.LogError(StatusCodes.Status403Forbidden, $"User {personalGoalNoteDetail.CreatedBy} is forbidden to perform this operation.");
                    return this.StatusCode(StatusCodes.Status403Forbidden, $"User {personalGoalNoteDetail.CreatedBy} is forbidden to perform this operation.");
                }
            }

            this.logger.LogInformation("Personal goal note details received are valid.");
            return this.Ok("Personal goal note details received are valid.");
        }
    }
}