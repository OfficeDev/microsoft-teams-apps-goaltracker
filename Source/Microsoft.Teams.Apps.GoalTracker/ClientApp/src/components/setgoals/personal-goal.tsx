// <copyright file="personal-goal.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import React from 'react';
import { Input, Loader } from '@fluentui/react-northstar';
import { createBrowserHistory } from "history";
import { IPersonalGoalDetail, IAddNewGoal } from "../../models/type";
import "../../styles/style.css";
import { SeverityLevel } from "@microsoft/applicationinsights-web";
import * as microsoftTeams from "@microsoft/teams-js";
import { CloseIcon } from '@fluentui/react-icons-northstar';
import Constants from "../../constants";
import SetGoal from './set-goal'
import { handleError } from '../../helpers/goal-helper'
import { savePersonalGoalDetails, getPersonalGoalDetails } from '../../api/personal-goal-api'
import { getApplicationInsightsInstance } from "../../helpers/app-insights";
import { Guid } from "guid-typescript";
import { WithTranslation, withTranslation } from "react-i18next";
import { TFunction } from "i18next";
import moment from 'moment';

interface IPersonalGoalState {
    goalName: string,
    startDate: string,
    minStartDate: string,
    endDate: string,
    endDateUTC: string,
    isReminderActive: boolean,
    reminderFrequency: number,
    personalGoals: IPersonalGoalDetail[],
    addNewGoalDetails: IAddNewGoal[],
    errorMessage: string,
    loading: boolean,
    isSaveButtonLoading: boolean,
    isSaveButtonDisabled: boolean,
    showError: boolean;
    screenWidth: number;
}

const browserHistory = createBrowserHistory({ basename: "" });

class PersonalGoal extends React.Component<WithTranslation, IPersonalGoalState>
{
    localize: TFunction;
    telemetry?: any = null;
    scope?: string | null;
    userAadObjectId?: string | null = null;
    userPrincipalName?: string | null = null;
    serviceURL: string | null = null;
    conversationId?: string | null = null;
    appInsights: any;
    goalCycleId?: string | null;
    theme?: string | null;

    constructor(props: any) {
        super(props);
        this.localize = this.props.t;
        this.state = {
            goalName: "",
            minStartDate: "",
            startDate: "",
            endDate: "",
            endDateUTC: "",
            isReminderActive: true,
            reminderFrequency: 0,
            personalGoals: [],
            addNewGoalDetails: [],
            errorMessage: "",
            loading: false,
            isSaveButtonLoading: false,
            isSaveButtonDisabled: false,
            showError: false,
            screenWidth: 0,
        };
        let search = window.location.search;
        let params = new URLSearchParams(search);
        this.serviceURL = params.get("serviceURL");
        this.telemetry = params.get("telemetry");
    }

    /** Called once component is mounted. */
    async componentDidMount() {
        microsoftTeams.initialize();
        microsoftTeams.getContext((context) => {
            this.userAadObjectId = context.userObjectId;
            this.userPrincipalName = context.userPrincipalName ? context.userPrincipalName : "";
            this.appInsights = getApplicationInsightsInstance(this.telemetry, browserHistory);
            this.getPersonalGoalDetails();
            this.theme = context.theme!;
            window.addEventListener("resize", this.update.bind(this));
            this.update();
        });
    }

    /**
    * get screen width real time
    */
    update = () => {
        if (window.innerWidth !== this.state.screenWidth) {
            this.setState({ screenWidth: window.innerWidth });
        }
    };

    /**
    *  Gets called when user changes send reminder check box.
    * */
    private setIsReminderActive = (e: any, checkboxProps: any) => {
        this.setState({
            isReminderActive: checkboxProps.checked
        });
    };

    /**
    *  Gets called when user changes reminder duration i.e weekly, bi-weekly, monthly, quarterly.
    * */
    private setReminder = (e, props) => {
        this.setState({
            reminderFrequency: props.value,
        });
    };

    /**
   *  Gets called when user changes goal name in text box.
   * */
    private addGoalFromTextBox = (event) => {
        this.setState({ goalName: event.target.value });
    };

    /**
    *  Gets called when user changes start date.
    * */
    private getStartDate = (startDate: Date | undefined) => {
        this.setState({ startDate: startDate?.toString()! });
        if (startDate) {
            this.setState({ showError: false, isSaveButtonDisabled: false, errorMessage: "" })
        }
    }

    /**
    *  Gets called when user changes end date.
    * */
    private getEndDate = (endDate: Date | undefined) => {
        this.setState({ endDateUTC: moment(endDate?.toUTCString()!).format(Constants.utcDateFormat), endDate: endDate?.toString()! });
        if (endDate) {
            this.setState({ showError: false, isSaveButtonDisabled: false, errorMessage: "" })
        }
    }

    /**
    * Validate personal goals on click of save button.
    * */
    private validateGoals() {
        let errorMessage: string = "";
        if (!errorMessage) {
            if (!this.state.startDate) {
                errorMessage = this.localize("emptyStartDateValidationText")
            }
            else if (!this.state.endDate) {
                errorMessage = this.localize("emptyEndDateValidationText");
            }
            else if (moment(new Date(this.state.endDate)) < moment(new Date(this.state.startDate)).add(30, 'd')) {
                errorMessage = this.localize("endDateValidationText")
            }
            else if (this.state.addNewGoalDetails.length === 0) {
                errorMessage = this.localize("emptyGoalNameValidationText")
            }
            else if (this.state.addNewGoalDetails.length) {
                this.state.addNewGoalDetails.forEach(goal => {
                    if (goal.goalName) {
                        goal.goalName = goal.goalName.trim();
                    }
                    if (goal.goalName === "") {
                        errorMessage = this.localize("emptyGoalNameValidationText")
                    }
                });
            }
        }
        if (errorMessage) {
            this.setState({ showError: true, errorMessage: errorMessage, isSaveButtonLoading: false, isSaveButtonDisabled: true });
            return false;
        }
        else {
            this.setState({ showError: false, errorMessage: "" });
            return true;
        }
    }

    /**
    * Gets called when user enters goal name.
    * */
    private goalNameChange = (goalId: string, event: any) => {
        if (event.target.value !== "") {
            this.setState({ showError: false, errorMessage: "", isSaveButtonDisabled: false });
        }
        this.appInsights.trackTrace({ message: `'goalNameChange' - Request initiated`, severityLevel: SeverityLevel.Information });
        let newGoalDetails = this.state.addNewGoalDetails;
        let personalGoals = this.state.personalGoals;
        let newGoalIndex = newGoalDetails.findIndex(goal => goal.key === goalId);
        // Update goal name in list UI
        newGoalDetails[newGoalIndex].goalName = event.target.value;
        newGoalDetails[newGoalIndex].header = <Input fluid className="add-goals-input" icon={<CloseIcon outline className="remove-goal-button " aria-label={this.localize("removeGoalIcon")} title="Close" onClick={event => this.removeGoals(goalId)} />} placeholder={this.localize("addGoalPlaceHolder")} value={newGoalDetails[newGoalIndex].goalName} title={newGoalDetails[newGoalIndex].goalName}  maxLength={Constants.maxAllowedGoalName} onChange={event => this.goalNameChange(goalId, event)} />
        this.setState({ addNewGoalDetails: newGoalDetails });
        // Update the goal name in personal goal array which is then saved to storage.
        if (personalGoals.length > 0 && personalGoals.some(goal => goal.PersonalGoalId === goalId)) {
            let personalGoalIndex = personalGoals.findIndex(goal => goal.PersonalGoalId === goalId);
            personalGoals[personalGoalIndex].GoalName = newGoalDetails[newGoalIndex].goalName;
        }
    };

    /**
    * Gets the personal goal details from storage.
    * */
    private getPersonalGoalDetails = async () => {
        this.appInsights.trackTrace({ message: `'getPersonalGoalDetails' - Request initiated`, severityLevel: SeverityLevel.Information });
        let addGoalDetails = this.state.addNewGoalDetails;
        const personalGoalDetailsResponse = await getPersonalGoalDetails();
        if (personalGoalDetailsResponse.data.length > 0) {
            this.setState({
                personalGoals: personalGoalDetailsResponse.data as IPersonalGoalDetail[]
            });

            this.state.personalGoals.forEach((personalGoal) => {
                addGoalDetails.push({
                    key: personalGoal.PersonalGoalId,
                    header: <Input fluid className="add-goals-input" icon={<CloseIcon outline className="remove-goal-button " aria-label={this.localize("removeGoalIcon")} title="Close" onClick={event => this.removeGoals(personalGoal.PersonalGoalId)} />} aria-label={this.localize("addGoalPlaceHolder")} placeholder={this.localize("addGoalPlaceHolder")} value={personalGoal.GoalName} title={personalGoal.GoalName} maxLength={Constants.maxAllowedGoalName} onChange={event => this.goalNameChange(personalGoal.PersonalGoalId, event)} />,
                    goalName: personalGoal.GoalName,
                });
            });
            let startDate = moment(this.state.personalGoals[0].StartDate).format(Constants.dateComparisonFormat);
            let todaysDate = moment(new Date().toDateString()).format(Constants.dateComparisonFormat);

            this.setState({
                startDate: this.state.personalGoals[0].StartDate,
                minStartDate: startDate > todaysDate
                    ? new Date().toDateString() : this.state.personalGoals[0].StartDate,
                endDate: this.state.personalGoals[0].EndDate,
                endDateUTC: this.state.personalGoals[0].EndDateUTC,
                reminderFrequency: this.state.personalGoals[0].ReminderFrequency,
                isReminderActive: this.state.personalGoals[0].IsReminderActive
            });
        };
    };

    /**
    * Gets called when user clicks on save button to save goals.
    * */
    private saveGoals = async () => {
        await this.setState({ isSaveButtonLoading: true, isSaveButtonDisabled: true })
        this.appInsights.trackTrace({ message: `'saveGoals' - Click on the save button`, severityLevel: SeverityLevel.Information });
        if (this.state.goalName) {
            let goalId = Guid.create().toString();
            let goalName = this.state.goalName;
            this.state.addNewGoalDetails.push({
                key: goalId,
                header: <Input fluid className="add-goals-input" icon={<CloseIcon outline className="remove-goal-button " aria-label={this.localize("removeGoalIcon")} title="Close" onClick={event => this.removeGoals(goalId)} />} aria-label={this.localize("addGoalPlaceHolder")} placeholder={this.localize("addGoalPlaceHolder")} value={goalName} title={goalName} maxLength={Constants.maxAllowedGoalName} onChange={event => this.goalNameChange(goalId, event)} />,
                goalName: goalName
            });

            this.setState({ goalName : ""});
        }

        if (this.validateGoals()) {
            let newGoalDetails = this.state.addNewGoalDetails;
            let personalGoals = this.state.personalGoals;
            this.setState({ showError: false, errorMessage: "" });

            // Edit goal scenario
            if (personalGoals.length > 0) {
                personalGoals.forEach((goal) => {
                    goal.UserAadObjectId = this.userAadObjectId;
                    goal.StartDate = moment(this.state.startDate.toString()).format(Constants.dateTimeOffsetFormat);
                    goal.EndDate = moment(this.state.endDate.toString()).format(Constants.dateTimeOffsetFormat);
                    goal.ServiceURL = this.serviceURL;
                    goal.ReminderFrequency = this.state.reminderFrequency;
                    goal.IsReminderActive = this.state.isReminderActive;
                    goal.EndDateUTC = moment(new Date(goal.EndDate!).toUTCString()).format(Constants.utcDateFormat);
                });
                this.goalCycleId = personalGoals[0].GoalCycleId;
            }

            let goalCycleId = this.goalCycleId ? this.goalCycleId : Guid.create().toString();
            newGoalDetails.forEach((goal) => {
                // Add new goal
                if (!personalGoals.some(personalGoal => personalGoal.PersonalGoalId === goal.key)) {
                    personalGoals.push({
                        UserAadObjectId: this.userAadObjectId,
                        PersonalGoalId: goal.key,
                        CreatedOn: new Date().toUTCString(),
                        CreatedBy: this.userPrincipalName,
                        LastModifiedOn: new Date().toUTCString(),
                        LastModifiedBy: this.userPrincipalName,
                        AdaptiveCardActivityId: "",
                        IsActive: true,
                        IsAligned: false,
                        IsDeleted: false,
                        ConversationId: this.conversationId,
                        GoalName: goal.goalName,
                        Status: 0,
                        StartDate: moment(this.state.startDate.toString()).format(Constants.dateTimeOffsetFormat),
                        EndDate: moment(this.state.endDate.toString()).format(Constants.dateTimeOffsetFormat),
                        ServiceURL: this.serviceURL,
                        ReminderFrequency: this.state.reminderFrequency,
                        IsReminderActive: this.state.isReminderActive,
                        TeamId: null,
                        TeamGoalName: null,
                        TeamGoalId: null,
                        EndDateUTC: moment(new Date(this.state.endDateUTC).toUTCString()).format(Constants.utcDateFormat),
                        NotesCount: 0,
                        GoalCycleId: goalCycleId,
                    })
                }
            });

            // Store goal details in table storage.
            let response = await this.saveGoalDetails();
            if (response) {
                let userAadObjectId = this.userAadObjectId;
                let activityId = this.state.personalGoals[0].AdaptiveCardActivityId;
                let goalCycleId = this.state.personalGoals[0].GoalCycleId;
                let command = activityId ? Constants.editPersonalGoal : Constants.setPersonalGoal;
                let toBot = { AdaptiveActionType: command, PersonalGoalDetails: this.state.personalGoals, UserAadObjectId: userAadObjectId, ActivityCardId: activityId, GoalCycleId: goalCycleId };
                microsoftTeams.getContext((context) => {
                    microsoftTeams.tasks.submitTask(toBot);
                });
            }
        }
    };

    /**
    *  Stores personal goal details in table storage.
    * */
    private saveGoalDetails = async () => {
        this.appInsights.trackTrace({ message: `'saveGoalDetails' - Request initiated`, severityLevel: SeverityLevel.Information });
        if (this.state.personalGoals.length > 0) {
            const saveGoalDetailsResponse = await savePersonalGoalDetails(this.state.personalGoals)
            if (saveGoalDetailsResponse.status !== 200 && saveGoalDetailsResponse.status !== 204) {
                this.setState({ isSaveButtonLoading: false, errorMessage: this.state.errorMessage, isSaveButtonDisabled: false });
                handleError(saveGoalDetailsResponse);
                return false;
            }
            this.appInsights.trackTrace({ message: `'saveGoalDetails' - Personal goal details saved and userAadObjectId=${this.userAadObjectId}`, severityLevel: SeverityLevel.Information });
            return true;
        }
    }

    /**
    *  Gets called when user clicks on close icon to remove goal.
    * */
    private removeGoals = async (goalId: string) => {
        this.appInsights.trackTrace({ message: `'remove' - close icon is clicked to remove goal`, severityLevel: SeverityLevel.Information });
        let removeGoal = this.state.addNewGoalDetails;
        let personalGoals = this.state.personalGoals;
        let index = removeGoal.findIndex(goal => goal.key === goalId);
        removeGoal.splice(index, 1);
        this.setState({ addNewGoalDetails: removeGoal });
        if (personalGoals.length > 0 && personalGoals.some(goal => goal.PersonalGoalId === goalId)) {
            let personalGoalIndex = personalGoals.findIndex(goal => goal.PersonalGoalId === goalId);
            personalGoals[personalGoalIndex].IsDeleted = true;
            personalGoals[personalGoalIndex].IsActive = false;
            this.setState({ personalGoals: personalGoals });
        }
    };

    /**
    *  Gets called when user clicks on add new goal button.
    * */
    private addNewGoal = (event) => {
        let errorText: string = "";
        let isSaveButtonDisabled = false;
        this.appInsights.trackTrace({ message: `'addNewGoal' - Request initiated`, severityLevel: SeverityLevel.Information });
        // validate if goal added are more than 15.
        if (this.state.addNewGoalDetails.length >= Constants.maxAllowedGoals) {
            errorText = this.localize("goalCountValidationText");
            isSaveButtonDisabled = true;
        }
        else {
            if (this.state.goalName === "") {
                errorText = this.localize("emptyGoalNameValidationText");
                isSaveButtonDisabled = true;
            }
            else {
                let addGoal = this.state.addNewGoalDetails;
                let goalId = Guid.create().toString();
                let goalName = this.state.goalName;
                this.setState({ goalName: "" });
                addGoal.push({
                    key: goalId,
                    header: <Input fluid className="add-goals-input" icon={<CloseIcon outline className="remove-goal-button " aria-label={this.localize("removeGoalIcon")} title="Close" onClick={event => this.removeGoals(goalId)} />} aria-label={this.localize("addGoalPlaceHolder")} placeholder={this.localize("addGoalPlaceHolder")} value={goalName} title={goalName} maxLength={Constants.maxAllowedGoalName} onChange={event => this.goalNameChange(goalId, event)} />,
                    goalName: goalName
                });
                this.setState({ addNewGoalDetails: addGoal, goalName: "" });
            }
        }
        this.setState({ errorMessage: errorText, isSaveButtonDisabled: isSaveButtonDisabled });
    };

    /**
    *  Renders set goal UI.
    * */
    render() {
        let contents = this.state.loading
            ? <p><em><Loader /></em></p>
            : <SetGoal
                errorMessage={this.state.errorMessage}
                goals={this.state.addNewGoalDetails}
                isReminderActive={this.state.isReminderActive}
                setIsReminderActive={this.setIsReminderActive}
                startDate={this.state.startDate}
                minStartDate={this.state.minStartDate}
                endDate={this.state.endDate}
                reminderFrequency={this.state.reminderFrequency}
                setReminder={this.setReminder}
                saveGoals={this.saveGoals}
                removeGoals={this.removeGoals}
                addNewGoal={this.addNewGoal}
                addGoalFromTextBox={this.addGoalFromTextBox}
                getStartDate={this.getStartDate}
                getEndDate={this.getEndDate}
                goalName={this.state.goalName}
                isSaveButtonLoading={this.state.isSaveButtonLoading}
                isSaveButtonDisabled={this.state.isSaveButtonDisabled}
                theme={this.theme}
                screenWidth={this.state.screenWidth}
            />
        return (
            <div className="container-div">
                {contents}
            </div>
        )
    }
}
export default withTranslation()(PersonalGoal);