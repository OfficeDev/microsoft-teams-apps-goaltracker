// <copyright file="edit-goal-detail.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import * as React from 'react';
import * as microsoftTeams from "@microsoft/teams-js";
import { Loader, Flex, Text, Input, Dropdown, Button, Checkbox, RadioGroup } from "@fluentui/react-northstar";
import { WithTranslation, withTranslation } from "react-i18next";
import { TFunction } from "i18next";
import { SeverityLevel } from "@microsoft/applicationinsights-web";
import { createBrowserHistory } from "history";
import { getApplicationInsightsInstance } from "../../helpers/app-insights";
import { getPersonalGoalDetailByGoalIdAsync, updatePersonalGoalDetail } from "../../api/personal-goal-api";
import { getPersonalGoalNoteDetails, savePersonalGoalNoteDetails, deletePersonalGoalNoteDetails } from "../../api/personal-goal-note-api";
import { getTeamGoalDetailByTeamGoalId } from "../../api/team-goal-api";
import { IPersonalGoalDetail, IPersonalGoalNoteDetail } from "../../models/type";
import { handleError, getGoalStatusCollection } from "../../helpers/goal-helper";
import PersonalGoalNote from "./personal-goal-notes"
import Constants from "../../constants";
let moment = require('moment');

interface IEditGoalState {
    loader: boolean,
    errorMessage: string,
    personalGoalDetail: IPersonalGoalDetail,
    notesData: IPersonalGoalNoteDetail[],
    deletedNotesData: IPersonalGoalNoteDetail[],
    goalStatus: any,
    isGoalDetailsLoading: boolean,
    isGoalSaved: boolean
}

const browserHistory = createBrowserHistory({ basename: "" });

/** Component for displaying edit personal goal. */
class EditGoal extends React.Component<WithTranslation, IEditGoalState> {
    localize: TFunction;
    telemetry?: any = null;
    appInsights: any;
    personalGoalId: string | null = null;
    userEmail?: any = null;
    teamId?: string | null;
    botId: string;
    appBaseUrl: string;
    goalCycle: string;
    notesCount: string;
    goalStatusCollection: any;

    constructor(props: any) {
        super(props);
        this.localize = this.props.t;
        this.state = {
            loader: false,
            errorMessage: "",
            personalGoalDetail: {} as IPersonalGoalDetail,
            notesData: [] as IPersonalGoalNoteDetail[],
            deletedNotesData: [] as IPersonalGoalNoteDetail[],
            goalStatus: {},
            isGoalDetailsLoading: false,
            isGoalSaved : false
        };

        let search = window.location.search;
        let params = new URLSearchParams(search);
        this.personalGoalId = params.get("goalId");
        this.botId = "";
        this.goalCycle = "";
        this.notesCount = "";
        this.appBaseUrl = window.location.origin;
        this.goalStatusCollection = getGoalStatusCollection(this.localize);
    }

    /**
    *  Called once component is mounted.
    * */
    async componentDidMount() {
        microsoftTeams.initialize();
        microsoftTeams.getContext((context) => {
            this.userEmail = context.upn;
            this.appInsights = getApplicationInsightsInstance(this.telemetry, browserHistory);
            this.getPersonalGoalDetails();
            this.getPersonalGoalNoteDetails();
        });
    }

    /** 
    *  Get personal goal details from storage.
    * */
    getPersonalGoalDetails = async () => {
        this.appInsights.trackTrace({ message: `'getPersonalGoalDetails' - Request initiated`, severityLevel: SeverityLevel.Information });
        this.setState({ loader: true });
        const personalGoalDetailsResponse = await getPersonalGoalDetailByGoalIdAsync(this.personalGoalId!);
        if (personalGoalDetailsResponse) {
            if (personalGoalDetailsResponse.status === 200) {
                let goalStatus = this.goalStatusCollection.find(goalStatus => goalStatus.value === personalGoalDetailsResponse.data.Status);
                this.setState({ personalGoalDetail: personalGoalDetailsResponse.data, goalStatus: goalStatus });

                if (personalGoalDetailsResponse.data.IsAligned) {
                    let teamGoalDetail = await this.getTeamGoalDetail();
                    this.goalCycle = `${moment(teamGoalDetail?.TeamGoalStartDate).format(Constants.goalCycleDateTimeFormat)} ${this.localize("goalCycleToText")} ${moment(teamGoalDetail?.TeamGoalEndDate).format(Constants.goalCycleDateTimeFormat)}`;
                }
                else {
                    this.goalCycle = `${moment(personalGoalDetailsResponse.data.StartDate).format(Constants.goalCycleDateTimeFormat)} ${this.localize("goalCycleToText")} ${moment(personalGoalDetailsResponse.data.EndDate).format(Constants.goalCycleDateTimeFormat)}`;
                }
            }
            else {
                handleError(personalGoalDetailsResponse);
            }

            this.appInsights.trackTrace({ message: `'getPersonalGoalDetails' - Request completed`, severityLevel: SeverityLevel.Information });
        }
        this.setState({ loader: false });
    }

    /**
    *  Get detail of team goal by specific team goal id.
    * */
    getTeamGoalDetail = async () => {
        this.appInsights.trackTrace({ message: `'getTeamGoalDetail' - Request initiated`, severityLevel: SeverityLevel.Information });
        this.setState({ loader: true });
        let teamGoalId = this.state.personalGoalDetail.TeamGoalId?.split(",")[0];
        const teamGoalDetailResponse = await getTeamGoalDetailByTeamGoalId(teamGoalId, this.state.personalGoalDetail.TeamId);
        if (teamGoalDetailResponse) {
            if (teamGoalDetailResponse.status === 200) {
                return teamGoalDetailResponse.data;
            }
            else {
                handleError(teamGoalDetailResponse);
            }

            this.appInsights.trackTrace({ message: `'getTeamGoalDetail' - Request completed`, severityLevel: SeverityLevel.Information });
        }
        this.setState({ loader: false });
    }

    /** 
    *  Get details of personal goal notes.
    * */
    getPersonalGoalNoteDetails = async () => {
        this.appInsights.trackTrace({ message: `'getPersonalGoalNoteDetails' - Request initiated`, severityLevel: SeverityLevel.Information });
        this.setState({ loader: true });
        const personalGoalNoteDetailsResponse = await getPersonalGoalNoteDetails(this.personalGoalId!);
        if (personalGoalNoteDetailsResponse) {
            if (personalGoalNoteDetailsResponse.status === 200) {
                let personalGoalNoteDetails: IPersonalGoalNoteDetail[] = personalGoalNoteDetailsResponse.data;
                personalGoalNoteDetails.forEach((goalNoteDetail) => {
                    goalNoteDetail.IsEdited = false;
                });
                this.setState({ notesData: personalGoalNoteDetails });
            }
            else {
                handleError(personalGoalNoteDetailsResponse);
            }

            this.appInsights.trackTrace({ message: `'getPersonalGoalNoteDetails' - Request completed`, severityLevel: SeverityLevel.Information });
        }
        this.setState({ loader: false });
    }

    /** 
    *  Save details of personal goal. 
    * */
    updatePersonalGoalDetail = async () => {
        this.appInsights.trackTrace({ message: `'updatePersonalGoalDetail' - Request initiated`, severityLevel: SeverityLevel.Information });
        this.setState({ isGoalDetailsLoading: true });

        if (this.validatePersonalGoalDetails()) {
            const personalGoalDetailsResponse = await updatePersonalGoalDetail(this.state.personalGoalDetail);
            if (personalGoalDetailsResponse) {
                if (personalGoalDetailsResponse.status === 200 || personalGoalDetailsResponse.status === 204) {
                    let personalGoalNoteDetailsResponse = true;
                    personalGoalNoteDetailsResponse = await this.savePersonalGoalNoteDetails();

                    if (personalGoalNoteDetailsResponse) {
                        this.setState({ errorMessage: "" });
                        microsoftTeams.getContext((context) => {
                            this.setState({ isGoalSaved: true});
                            microsoftTeams.tasks.submitTask();
                        });
                    }
                    else {
                        this.setState({ isGoalDetailsLoading: false, errorMessage: this.localize("goalNoteDetailsSubmitError") });
                    }
                }
                else {
                    this.setState({ isGoalDetailsLoading: false, errorMessage: this.localize("goalDetailsSubmitError") });
                }

                this.appInsights.trackTrace({ message: `'updatePersonalGoalDetail' - Request completed`, severityLevel: SeverityLevel.Information });
            }
        }
    }

    /** 
    *  Save details of personal goal notes. 
    * */
    savePersonalGoalNoteDetails = async () => {
        this.appInsights.trackTrace({ message: `'savePersonalGoalNoteDetails' - Request initiated`, severityLevel: SeverityLevel.Information });
        if (this.state.notesData.length > 0) {
            const personalGoalNoteDetailsResponse = await savePersonalGoalNoteDetails(this.state.notesData);
            if (personalGoalNoteDetailsResponse) {
                if (personalGoalNoteDetailsResponse.status !== 200 && personalGoalNoteDetailsResponse.status !== 204) {
                    return false;
                }
            }
        }

        if (this.state.deletedNotesData.length > 0) {
            let deletedNotesData = this.state.deletedNotesData;
            let deletedNotesIds: string[] = [];

            deletedNotesData.forEach((deletedNote) => {
                deletedNotesIds.push(deletedNote.PersonalGoalNoteId);
            });

            const personalGoalNoteDetailsResponse = await deletePersonalGoalNoteDetails(deletedNotesIds);
            if (personalGoalNoteDetailsResponse.status !== 200 && personalGoalNoteDetailsResponse.status !== 204) {
                return false;
            }
        }

        this.appInsights.trackTrace({ message: `'savePersonalGoalNoteDetails' - Request completed`, severityLevel: SeverityLevel.Information });
        return true;
    }

    /**
    *  Validate details of personal goals.
    * */
    validatePersonalGoalDetails = () => {
        if (!this.state.personalGoalDetail.GoalName) {
            this.setState({ errorMessage: this.localize("goalNameError"), isGoalDetailsLoading: false })
            return false;
        }

        if (this.state.notesData && this.state.notesData.length > 0) {
            let notesData = this.state.notesData;
            let personalGoalNote = notesData.find(personalGoalNote => !personalGoalNote.PersonalGoalNoteDescription);

            if (personalGoalNote) {
                this.setState({ errorMessage: this.localize("goalNoteError"), isGoalDetailsLoading: false })
                return false;
            }
        }

        this.setState({ errorMessage: ""})
        return true;
    }

    /**
    *   Handles goal status drop down change.
    */
    handleGoalStatusChange = (event: any, dropdownProps?: any) => {
        let goalsData = this.state.personalGoalDetail;
        let selectedStatus = dropdownProps.value;
        goalsData.Status = selectedStatus.value;
        this.setState({ personalGoalDetail: goalsData, goalStatus : selectedStatus });
    }

    /**
    *   Handles goal name change.
    */
    handleGoalNameChange = (event) => {
        let goalsData = this.state.personalGoalDetail;
        goalsData.GoalName = event.target.value;
        this.setState({ personalGoalDetail: goalsData });
    }

    /**
    *   Handles goal note delete button click.
    */
    onGoalNoteDeleteButtonClick = (goalNoteId: string) => {
        let notesDetails = this.state.notesData;
        let deletedNoteDetails = notesDetails.filter(noteDetails => noteDetails.PersonalGoalNoteId === goalNoteId)[0];  
        this.state.deletedNotesData.push(deletedNoteDetails);

        notesDetails = notesDetails.filter(noteDetails => noteDetails.PersonalGoalNoteId !== goalNoteId);
        this.setState({ notesData: notesDetails });
    }

    /**
    *   Handles goal note edit button click.
    */
    onGoalNoteEditButtonClick = (index: number, event: any) => {
        let noteDetails = this.state.notesData;
        noteDetails[index].IsEdited = true;
        this.setState({ notesData: noteDetails });
    }

    /**
    *   Handles goal note text change.
    */
    handleGoalNoteChange = (index: number, event: any) => {
        let noteDetails = this.state.notesData;
        noteDetails[index].PersonalGoalNoteDescription = event.target.value;
        this.setState({ notesData: noteDetails });
    }

    /**
    *   Get wrapper for page which acts as container for all child components
    */
    private getGoalDetails = () => {
        if (this.state.loader) {
            return (
                <div className="loader">
                    <Loader />
                </div>
            );
        }
        else {
            return (
                <>
                    <div className="edit-goal-details-container">
                        <Flex gap="gap.small" >
                            {this.goalCycle && <Text weight="bold" className="control-spacing goal-cycle-text" align="center" content={`${this.localize("goalCycleText")}: ${this.goalCycle}`} />}
                        </Flex>
                        <Flex gap="gap.large" vAlign="center" className="control-spacing edit-goal-title" >
                            <Text align="center" content={this.localize("goalNameHeader")} />
                        </Flex>
                        <Flex gap="gap.large" vAlign="center" className="control-padding">
                            <Input fluid className="goal-name-input" value={this.state.personalGoalDetail.GoalName} title={this.state.personalGoalDetail.GoalName} onChange={this.handleGoalNameChange} maxLength={Constants.maxAllowedGoalName} />
                        </Flex>
                        <Flex gap="gap.large" vAlign="center" className="control-spacing edit-goal-title" >
                            <Text align="center" content={this.localize("goalStatusText")} />
                        </Flex>
                        <Flex gap="gap.large" vAlign="center" className="control-padding">
                            <Dropdown fluid className="width-small" items={this.goalStatusCollection} value={this.state.goalStatus} onChange={this.handleGoalStatusChange} />
                        </Flex>
                        
                        <PersonalGoalNote
                            notesData={this.state.notesData}
                            onGoalNoteDeleteButtonClick={this.onGoalNoteDeleteButtonClick}
                            onGoalNoteEditButtonClick={this.onGoalNoteEditButtonClick}
                            handleGoalNoteChange={this.handleGoalNoteChange}
                        />
                        
                    </div>
                    <div className="footer">
                        <div className="error">
                            <Flex gap="gap.small">
                                {this.state.errorMessage !== null && <Text className="small-margin-left" content={this.state.errorMessage} error />}
                            </Flex>
                        </div>
                        <div className="save-button">
                            <Flex gap="gap.smaller">
                                <Button primary content={this.localize("saveButtonText")} loading={this.state.isGoalDetailsLoading} disabled={this.state.isGoalDetailsLoading} onClick={this.updatePersonalGoalDetail} />
                            </Flex>
                        </div>
                    </div>
                </>
            );
        }
    }

   /**
   * Renders the component
   */
    public render() {
        return (
            <div>
                {this.getGoalDetails()}
            </div>
        );
    }
}

export default withTranslation()(EditGoal);