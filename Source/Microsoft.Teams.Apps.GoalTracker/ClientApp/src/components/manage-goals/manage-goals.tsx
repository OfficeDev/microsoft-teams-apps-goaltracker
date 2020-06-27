// <copyright file="manage-goals.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import * as React from 'react';
import * as microsoftTeams from "@microsoft/teams-js";
import { Loader, Flex, Text } from "@fluentui/react-northstar";
import { WithTranslation, withTranslation } from "react-i18next";
import { TFunction } from "i18next";
import { SeverityLevel } from "@microsoft/applicationinsights-web";
import { createBrowserHistory } from "history";
import { getApplicationInsightsInstance } from "../../helpers/app-insights";
import PersonalGoalTable from "./personal-goals-table";
import { getPersonalGoalDetails, deletePersonalGoalDetail } from "../../api/personal-goal-api";
import { getTeamGoalDetailsByTeamId } from "../../api/team-goal-api";
import { getPersonalGoalNotesCount } from "../../api/personal-goal-note-api";
import { handleError, getGoalStatusCollection } from "../../helpers/goal-helper";
import Constants from "../../constants";
import { IPersonalGoalDetail } from "../../models/type";
let moment = require('moment');

interface IManageGoalState {
    loader: boolean,
    errorMessage: string,
    goalsData: IPersonalGoalDetail[],
    screenWidth: number,
}

const browserHistory = createBrowserHistory({ basename: "" });

/** Component for displaying personal goals tab. */
class ManageGoal extends React.Component<WithTranslation, IManageGoalState> {
    localize: TFunction;
    telemetry?: any = null;
    appInsights: any;
    teamId?: string | null;
    botId: string;
    appBaseUrl: string;
    appUrl: string = (new URL(window.location.href)).origin;
    goalCycle: string;

    constructor(props: any) {
        super(props);
        this.localize = this.props.t;
        this.state = {
            loader: false,
            errorMessage: "",
            goalsData: [],
            screenWidth: 0
        };

        this.botId = "";
        this.goalCycle = "";
        this.appBaseUrl = window.location.origin;
    }

    /** Called once component is mounted. */
    async componentDidMount() {
        microsoftTeams.initialize();
        microsoftTeams.getContext((context) => {
            this.appInsights = getApplicationInsightsInstance(this.telemetry, browserHistory);
            this.getPersonalGoalDetails();
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
    *  Get personal goal details from storage.
    * */
    getPersonalGoalDetails = async () => {
        this.appInsights.trackTrace({ message: `'getPersonalGoalDetails' - Request initiated`, severityLevel: SeverityLevel.Information });
        this.setState({ loader: true });
        const personalGoalDetailsResponse = await getPersonalGoalDetails();
        if (personalGoalDetailsResponse) {
            if (personalGoalDetailsResponse.status === 200) {
                let personalGoalDetails: any = personalGoalDetailsResponse;
                this.setState({ goalsData: personalGoalDetails.data });
                if (personalGoalDetails.data && personalGoalDetails.data.length > 0) {
                    this.goalCycle = `${moment(personalGoalDetails.data[0].StartDate).format(Constants.goalCycleDateTimeFormat)} ${this.localize("goalCycleToText")} ${moment(personalGoalDetails.data[0].EndDate).format(Constants.goalCycleDateTimeFormat)}`;
                    let goalDetails = personalGoalDetails.data.find(goalDetails => goalDetails.IsAligned);
                    if (goalDetails) {
                        await this.getTeamGoalDetails(goalDetails.TeamId);
                    }
                    else
                    {
                        personalGoalDetails.data.forEach((goalDetail) => {
                            goalDetail.TeamGoalName = this.localize("notAlignedTeamGoaltext");
                        });
                    }
                    await this.getPersonalGoalNoteDetails();
                }
            }
            else {
                handleError(personalGoalDetailsResponse);
            }
        }
        this.setState({ loader: false });
    }

    /** 
    *  Get team goal details from storage.
    * */
    getTeamGoalDetails = async (teamId: string) => {
        this.appInsights.trackTrace({ message: `'getTeamGoalDetails' - Request initiated`, severityLevel: SeverityLevel.Information });
        this.setState({ loader: true });
        const teamGoalDetailsResponse = await getTeamGoalDetailsByTeamId(teamId);
        if (teamGoalDetailsResponse) {
            if (teamGoalDetailsResponse.status === 200) {
                let teamGoalDetails: any = teamGoalDetailsResponse.data;
                let personalGoalDetails = this.state.goalsData;
                
                personalGoalDetails.forEach((goalDetail) => {
                    if (goalDetail.IsAligned) {
                        goalDetail.TeamGoalName = "";
                        goalDetail.TeamGoalId?.split(",").forEach((teamGoalId) => {
                            let alignedTeamGoalDetail = teamGoalDetails && teamGoalDetails.find(teamGoalDetail => teamGoalDetail.TeamGoalId === teamGoalId);
                            goalDetail.TeamGoalName += (alignedTeamGoalDetail && alignedTeamGoalDetail.TeamGoalName.trim() + ", ") || this.localize("notAlignedTeamGoaltext");
                        });
                        goalDetail.TeamGoalName = goalDetail.TeamGoalName?.trim().slice(0, -1);
                    }
                    else {
                        goalDetail.TeamGoalName = this.localize("notAlignedTeamGoaltext")
                    }
                });
                this.setState({ goalsData: personalGoalDetails });
            }
            else {
                handleError(teamGoalDetailsResponse);
            }
        }
        this.setState({ loader: false });
    }

    /** 
    *  Get details of personal goal 
    * */
    getPersonalGoalNoteDetails = async () => {
        this.appInsights.trackTrace({ message: `'getPersonalGoalNotesCount' - Request initiated`, severityLevel: SeverityLevel.Information });
        this.setState({ loader: true });
        const getPersonalGoalNotesCountResponse = await getPersonalGoalNotesCount();
        if (getPersonalGoalNotesCountResponse) {
            if (getPersonalGoalNotesCountResponse.status === 200) {
                let personalGoalNoteDetails: any = getPersonalGoalNotesCountResponse.data;
                let personalGoalDetails = this.state.goalsData;
                personalGoalDetails.forEach((goalDetail) => {
                    let personalGoalNoteDetail = personalGoalNoteDetails && personalGoalNoteDetails.find(personalGoalNoteDetail => personalGoalNoteDetail.personalGoalId === goalDetail.PersonalGoalId);
                    goalDetail.NotesCount = (personalGoalNoteDetail && personalGoalNoteDetail.notesCount) || 0;
                });

                this.setState({ goalsData: personalGoalDetails });
            }
            else {
                handleError(getPersonalGoalNotesCountResponse);
            }
        }
        this.setState({ loader: false });
    }

    /**
    *  Method deletes personal goal details from storage.
    * */
    deletePersonalGoalDetail = async (personalGoalDetail: IPersonalGoalDetail) => {
        this.setState({ loader: true });
        const deletePersonalGoalResponse = await deletePersonalGoalDetail(personalGoalDetail.PersonalGoalId);
        if (deletePersonalGoalResponse) {
            if (deletePersonalGoalResponse.status === 200) {
                let personalGoalDetails = this.state.goalsData;
                personalGoalDetails = personalGoalDetails.filter((goalDetail) => goalDetail.PersonalGoalId !== personalGoalDetail.PersonalGoalId);
                this.setState({ goalsData: personalGoalDetails });
                return true;
            }
            else {
                handleError(deletePersonalGoalResponse);
            }
        }
        this.setState({ loader: false });
        return false;
    }

    /**
    *  Gets called when user clicks on delete goal button.
    * */
    onDeleteButtonClick = (personalGoalDetail: IPersonalGoalDetail) => {
        personalGoalDetail.IsActive = false;
        personalGoalDetail.IsDeleted = true;
        this.deletePersonalGoalDetail(personalGoalDetail);
        this.setState({ loader: false });
        return true;
    };

    /**
    *  Handles task module submit event.
    * */
    submitHandler = async (err, result) => {
        this.appInsights.trackTrace(`Submit handler - err: ${err} - result: ${result}`);
        this.getPersonalGoalDetails();
    };

    /**
    *   Navigate to edit personal goal page.
    */
    onPersonalGoalClick = (personalGoalId: string, t:any) => {
        microsoftTeams.tasks.startTask({
            completionBotId: this.botId,
            title: t('editGoalDetailTitle'),
            height: 600,
            width: 600,
            url: `${this.appBaseUrl}/edit-goal-detail?goalId=${personalGoalId}`,
        }, this.submitHandler);
    }

    /**
    *   Renders goal cycle information.
    */
    private pageHeader = () => {
        return (
            <Flex gap="gap.small">
                {this.goalCycle && <Text weight="bold" className="goal-cycle" align="center" content={`${this.localize("goalCycleText")}: ${this.goalCycle}`} />}
            </Flex>
        );
    }

    /**
    *   Get wrapper for page which acts as container for all child components.
    */
    private getGoalDetails = () => {
        if (this.state.loader) {
            return (
                <div className="loader">
                    <Loader />
                </div>
            );
        }
        else if (this.state.goalsData && this.state.goalsData.length > 0) {
            return (
                <div>
                    <PersonalGoalTable
                        screenWidth={this.state.screenWidth}
                        goalsData={this.state.goalsData}
                        goalStatus={getGoalStatusCollection(this.localize)}
                        onDeleteButtonClick={this.onDeleteButtonClick}
                        onPersonalGoalClick={this.onPersonalGoalClick}
                    />
                </div>
            );
        }
        else {
            return (
                <Flex className="error-container" hAlign="center" vAlign="stretch">
                    <div>
                        <div><Text content={this.localize('noActiveGoalsMessage')} /></div>
                    </div>
                </Flex>
            )
        }
    }

   /**
   *    Renders the component.
   */
    public render() {
        return (
            <div className="container-tab" >
                <div className="accordian-container">
                    {this.pageHeader()}
                    <div>
                        {this.getGoalDetails()}
                    </div>
                </div>
            </div>
        );
    }
}

export default withTranslation()(ManageGoal);