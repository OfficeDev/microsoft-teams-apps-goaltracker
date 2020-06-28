// <copyright file="team-goal.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import React from 'react';
import { Input, Loader, Flex } from '@fluentui/react-northstar';
import { createBrowserHistory } from "history";
import { ITeamGoalDetail, IAddNewGoal } from "../../models/type";
import "../../styles/style.css";
import { SeverityLevel } from "@microsoft/applicationinsights-web";
import * as microsoftTeams from "@microsoft/teams-js";
import { CloseIcon } from '@fluentui/react-icons-northstar';
import Constants from "../../constants";
import SetGoal from './set-goal'
import { saveTeamGoalDetails, getTeamGoalDetailsByTeamId, getTeamOwnerDetails } from '../../api/team-goal-api'
import { Guid } from "guid-typescript";
import { WithTranslation, withTranslation } from "react-i18next";
import { TFunction } from "i18next";
import { getApplicationInsightsInstance } from '../../helpers/app-insights';
import moment from 'moment';

interface ITeamGoalState {
    goalName: string,
    startDate: string,
    minStartDate: string,
    endDate: string,
    endDateUTC: string,
    isReminderActive: boolean,
    reminderFrequency: number,
    teamGoals: ITeamGoalDetail[],
    addNewGoalDetails: IAddNewGoal[],
    errorMessage: string,
    loading: boolean,
    isSaveButtonLoading: boolean,
    isSaveButtonDisabled: boolean;
    showError: boolean;
    isTeamOwner: boolean;
    screenWidth: number;
}

const browserHistory = createBrowserHistory({ basename: "" });

class TeamGoal extends React.Component<WithTranslation, ITeamGoalState>
{
    localize: TFunction;
    telemetry?: any = null;
    scope?: string | null;
    teamId?: string | null = null;
    userAADObjectId?: string | null = null;
    userPrincipalName?: string | null = null;
    teamGroupId?: string | null = null;
    conversationId?: string | null = null;
    appInsights: any;
    serviceURL: string | null = null;
    goalCycleId?: string | null;
    theme?: string | null;

    constructor(props: any) {
        super(props);
        this.localize = this.props.t;
        this.state = {
            goalName: "",
            startDate: "",
            minStartDate: "",
            endDate: "",
            endDateUTC: "",
            isReminderActive: true,
            reminderFrequency: 0,
            teamGoals: [],
            addNewGoalDetails: [],
            errorMessage: "",
            loading: true,
            isSaveButtonLoading: false,
            isSaveButtonDisabled: false,
            showError: false,
            isTeamOwner: false,
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
        microsoftTeams.getContext(async (context) => {
            this.theme = context.theme!;
            this.teamId = context.teamId;
            this.teamGroupId = context.groupId;
            this.userAADObjectId = context.userObjectId;
            this.userPrincipalName = context.userPrincipalName ? context.userPrincipalName : "";
            this.appInsights = getApplicationInsightsInstance(this.telemetry, browserHistory);
            // Validate if user is team owner.

            var validationResponse = await this.validateIfTeamOwner();

            if (validationResponse) {
                this.getTeamGoalDetails(this.teamId);
            }
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
        if (endDate !== null || endDate !== "") {
            this.setState({ showError: false, isSaveButtonDisabled: false, errorMessage: "" })
        }
    }

    /**
    * Gets called when user changes goal name. 
    * */
    private goalNameChange = (goalId: string, event: any) => {
        if (event.target.value !== "") {
            this.setState({ showError: false, errorMessage: "", isSaveButtonDisabled: false });
        }
        this.appInsights.trackTrace({ message: `'goalNameChange' - Request initiated`, severityLevel: SeverityLevel.Information });
        let newGoals = this.state.addNewGoalDetails;
        let teamGoals = this.state.teamGoals;
        let newGoalIndex = newGoals.findIndex(goal => goal.key === goalId);
        newGoals[newGoalIndex].goalName = event.target.value;
        newGoals[newGoalIndex].header = <Input fluid className="add-goals-input" icon={<CloseIcon outline className="remove-goal-button " aria-label={this.localize("removeGoalIcon")} title="Close" onClick={event => this.removeGoals(goalId)} />} placeholder={this.localize("addGoalPlaceHolder")} value={newGoals[newGoalIndex].goalName} title={newGoals[newGoalIndex].goalName} maxLength={Constants.maxAllowedGoalName} onChange={event => this.goalNameChange(goalId, event)} />
        this.setState({ addNewGoalDetails: newGoals });

        // Update the goal name to team goal array that is saved to storage.
        if (teamGoals.length > 0 && teamGoals.some(goal => goal.TeamGoalId === goalId)) {
            let teamGoalIndex = teamGoals.findIndex(goal => goal.TeamGoalId === goalId);
            teamGoals[teamGoalIndex].TeamGoalName = newGoals[newGoalIndex].goalName;
        }
    };

    /**
    * Gets the team goal details from storage.
    * */
    private getTeamGoalDetails = async (teamId) => {
        this.appInsights.trackTrace({ message: `'getTeamGoalDetails' - Request initiated`, severityLevel: SeverityLevel.Information });
        let addGoalDetails = this.state.addNewGoalDetails;
        const teamGoalDetailsResponse = await getTeamGoalDetailsByTeamId(teamId);
        if (teamGoalDetailsResponse.data.length > 0) {
            this.setState({
                teamGoals: teamGoalDetailsResponse.data as ITeamGoalDetail[]
            });
            this.state.teamGoals.forEach((teamGoal) => {
                addGoalDetails.push({
                    key: teamGoal.TeamGoalId,
                    header: <Input fluid className="add-goals-input" icon={<CloseIcon outline className="remove-goal-button " aria-label={this.localize("removeGoalIcon")} title="Close" onClick={event => this.removeGoals(teamGoal.TeamGoalId)} />} aria-label={this.localize("addGoalPlaceHolder")} placeholder={this.localize("addGoalPlaceHolder")} value={teamGoal.TeamGoalName} title={teamGoal.TeamGoalName} maxLength={Constants.maxAllowedGoalName} onChange={event => this.goalNameChange(teamGoal.TeamGoalId, event)} />,
                    goalName: teamGoal.TeamGoalName,
                });
            });

            let startDate = moment(this.state.teamGoals[0].TeamGoalStartDate).format(Constants.dateComparisonFormat);
            let todaysDate = moment(new Date().toDateString()).format(Constants.dateComparisonFormat);

            this.setState({
                startDate: this.state.teamGoals[0].TeamGoalStartDate,
                minStartDate: startDate > todaysDate
                    ? new Date().toDateString() : this.state.teamGoals[0].TeamGoalStartDate,
                endDate: this.state.teamGoals[0].TeamGoalEndDate,
                reminderFrequency: this.state.teamGoals[0].ReminderFrequency,
                isReminderActive: this.state.teamGoals[0].IsReminderActive
            });
        }
        this.setState({ loading: false })
    };

    /**
    * Gets called when user clicks on save button to save goals.
    * */
    private saveGoals = async () => {

        await this.setState({ isSaveButtonLoading: true, isSaveButtonDisabled: true });
        this.appInsights.trackTrace({ message: `'saveGoals' - Save button is clicked`, severityLevel: SeverityLevel.Information });
        if (this.state.goalName) {
            let goalId = Guid.create().toString();
            let goalName = this.state.goalName;
            this.state.addNewGoalDetails.push({
                key: goalId,
                header: <Input fluid className="add-goals-input" icon={<CloseIcon outline className="remove-goal-button " aria-label={this.localize("removeGoalIcon")} title="Close" onClick={event => this.removeGoals(goalId)} />} aria-label={this.localize("addGoalPlaceHolder")} placeholder={this.localize("addGoalPlaceHolder")} value={goalName} title={goalName} maxLength={Constants.maxAllowedGoalName} onChange={event => this.goalNameChange(goalId, event)} />,
                goalName: goalName
            });

            this.setState({ goalName: "" });
        }

        if (this.validateGoals()) {
            let newGoalDetails = this.state.addNewGoalDetails;
            let teamGoals = this.state.teamGoals;
            this.setState({ showError: false, errorMessage: "" });

            // Edit goal scenario
            if (teamGoals.length > 0) {
                teamGoals.forEach((goal) => {
                    goal.TeamGoalStartDate = moment(this.state.startDate.toString()).format(Constants.dateTimeOffsetFormat);
                    goal.TeamGoalEndDate = moment(this.state.endDate.toString()).format(Constants.dateTimeOffsetFormat);
                    goal.ReminderFrequency = this.state.reminderFrequency;
                    goal.IsReminderActive = this.state.isReminderActive;
                    goal.TeamGoalEndDateUTC = this.state.endDateUTC;
                    goal.ServiceURL = this.serviceURL;
                });
                this.goalCycleId = teamGoals[0].GoalCycleId;
            }

            let goalCycleId = this.goalCycleId ? this.goalCycleId : Guid.create().toString();
            newGoalDetails.forEach((goal) => {
                // Adding new goals
                if (!teamGoals.some(teamGoal => teamGoal.TeamGoalId === goal.key)) {
                    teamGoals.push({
                        TeamGoalId: goal.key,
                        TeamGoalName: goal.goalName,
                        TeamGoalStartDate: moment(this.state.startDate.toString()).format(Constants.dateTimeOffsetFormat),
                        TeamGoalEndDate: moment(this.state.endDate.toString()).format(Constants.dateTimeOffsetFormat),
                        ReminderFrequency: this.state.reminderFrequency,
                        IsReminderActive: this.state.isReminderActive,
                        CreatedOn: new Date().toUTCString(),
                        CreatedBy: this.userPrincipalName,
                        LastModifiedOn: new Date().toUTCString(),
                        LastModifiedBy: null,
                        IsActive: true,
                        IsDeleted: false,
                        TeamId: this.teamId,
                        AdaptiveCardActivityId: "",
                        TeamGoalEndDateUTC: moment(new Date(this.state.endDateUTC).toUTCString()).format(Constants.utcDateFormat),
                        ServiceURL: this.serviceURL,
                        GoalCycleId: goalCycleId,
                    })
                }
            });

            // Store goal details in table storage.
            let response = await this.saveGoalDetails();
            if (response) {
                let teamId = this.teamId;
                let activityId = this.state.teamGoals[0].AdaptiveCardActivityId;
                let command = activityId ? Constants.editTeamGoal : Constants.setTeamGoal;
                let goalCycleId = this.state.teamGoals[0].GoalCycleId;
                let toBot = { AdaptiveActionType: command, TeamGoalDetails: this.state.teamGoals, TeamId: teamId, ActivityCardId: activityId, GoalCycleId: goalCycleId };
                microsoftTeams.getContext((context) => {
                    microsoftTeams.tasks.submitTask(toBot);
                });
            }
        }
    };

    /**
    * Validate team goals on click of save button.
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

            // End date must be 30 days more than star date.
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
            this.setState({ showError: true, errorMessage: errorMessage, isSaveButtonLoading: false, isSaveButtonDisabled: true, loading: false });
            return false;
        }
        else {
            this.setState({ showError: false, errorMessage: "" });
            return true;
        }
    }

    /**
    *  Stores goal details in table storage.
    * */
    private saveGoalDetails = async () => {
        this.appInsights.trackTrace({ message: `'saveGoalDetails' - Request initiated`, severityLevel: SeverityLevel.Information });
        if (this.state.teamGoals.length > 0) {
            const saveGoalDetailsResponse = await saveTeamGoalDetails(this.state.teamGoals, this.teamGroupId)
            if (saveGoalDetailsResponse.status !== 200 && saveGoalDetailsResponse.status !== 204) {
                this.setState({ isSaveButtonLoading: false, errorMessage: this.state.errorMessage, isSaveButtonDisabled: false });
                return false;
            }
            this.appInsights.trackTrace({ message: `'saveGoalDetails' - Team goal details saved and teamId=${this.teamId}`, severityLevel: SeverityLevel.Information });
            return true;
        }
    }

    /**
    *  Gets called when user clicks on close icon to remove goal.
    * */
    private removeGoals = async (goalId: string) => {
        this.appInsights.trackTrace({ message: `'remove' - close icon is clicked to remove goal`, severityLevel: SeverityLevel.Information });
        let removeGoal = this.state.addNewGoalDetails;
        let teamGoal = this.state.teamGoals;
        let index = removeGoal.findIndex(goal => goal.key === goalId);
        removeGoal.splice(index, 1);
        this.setState({ addNewGoalDetails: removeGoal });
        if (teamGoal.length > 0 && teamGoal.some(goal => goal.TeamGoalId === goalId)) {
            let teamGoalIndex = teamGoal.findIndex(goal => goal.TeamGoalId === goalId);
            teamGoal[teamGoalIndex].IsDeleted = true;
            teamGoal[teamGoalIndex].IsActive = false;
            this.setState({ teamGoals: teamGoal });
        }
    };

    /**
    *  Gets called when user clicks on add new goal button.
    * */
    private addNewGoal = async () => {
        let errorText: string = "";
        let isSaveButtonDisabled = false;
        this.appInsights.trackTrace({ message: `'addNewGoal' - Request initiated`, severityLevel: SeverityLevel.Information });
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

                // Add new goal row.
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
    *  Gets called at page load to check if user is a team owner.
    * */
    private validateIfTeamOwner = async () => {
        this.appInsights.trackTrace({ message: `'validateIfTeamOwner' - Request initiated`, severityLevel: SeverityLevel.Information });
        this.setState({
            loading: true
        });
        var teamOwnerDetailsResponse = await getTeamOwnerDetails(this.teamGroupId);
        if (teamOwnerDetailsResponse.status !== 200 && teamOwnerDetailsResponse.status !== 204) {
            this.appInsights.trackTrace({ message: `'validateIfTeamOwner' - Error while getting team owner details`, severityLevel: SeverityLevel.Information });
            await this.setState({
                isTeamOwner: false,
                loading: false
            });
            return false;
        }
        await this.setState({
            isTeamOwner: true
        });
        return true;
    }

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
        if (this.state.loading) {
            return (
                <p><em><Loader /></em></p>
            )
        }
        else {
            if (this.state.isTeamOwner) {
                return (
                    <div className="container-div">
                        {contents}
                    </div>
                )
            }
            else {
                return (
                    <div>
                        <Flex>
                            <Flex.Item size="size.half">
                                <div className="not-authorized-error">{this.localize("unauthorizedAccessMessage")}</div>
                            </Flex.Item>
                        </Flex>
                    </div >
                )
            }
        }
    }
}
export default withTranslation()(TeamGoal);