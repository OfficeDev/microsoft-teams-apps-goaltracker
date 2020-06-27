// <copyright file="align-goal.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import * as React from "react";
import { SeverityLevel } from "@microsoft/applicationinsights-web";
import * as microsoftTeams from "@microsoft/teams-js";
import moment from "moment";
import { WithTranslation, withTranslation } from "react-i18next";
import { TFunction } from "i18next";
import { Dropdown, Button, Loader, Flex, Divider, Text, Table, CloseIcon, TrashCanIcon } from "@fluentui/react-northstar";
import { createBrowserHistory } from "history";
import { getApplicationInsightsInstance } from "../../helpers/app-insights";
import { getTeamGoalDetailsByTeamId } from "../../api/team-goal-api";
import { getPersonalGoalDetails, savePersonalGoalDetails } from "../../api/personal-goal-api";
import { handleError } from "../../helpers/goal-helper";
import { IPersonalGoalDetail } from "../../models/type";
import AlignGoalSuccessScreen from './align-goal-success-screen'
import Constants from "../../constants";
import "../../styles/style.css";
import { Separator } from "office-ui-fabric-react";

interface ITeamGoalProps {
    key: string,
    header: string,
    TeamGoalId: string,
    TeamGoalName: string,
    TeamGoalStartDate: string,
    TeamGoalEndDate: string,
};

interface IPersonalGoalProps {
    key: string,
    header: string,
    GoalName: string,
    PersonalGoalId: string,
};

interface IState {
    loading: boolean,
    isAlignGoalLoading: boolean,
    isAlignToAddButtonDisabled: boolean,
    isAlignGoalButtonDisabled: boolean,
    isAlignGoalButtonLoading: boolean,
    isAlignedGoalsSubmitted: boolean,
    isSeeYourGoalButtonDisabled: boolean,
    isOkayButtonDisabled: boolean,
    isSeeYourGoalsButtonLoading: boolean,
    errorInAddToAlignGoal: string,
    personalGoalSelection: string,
    teamGoalSelection: string,
    teamGoalDetails: ITeamGoalProps[],
    personalGoalDetails: IPersonalGoalProps[],
    allPersonalGoalDetails: IPersonalGoalDetail[],
    alignGoalDetails: IPersonalGoalDetail[],
    screenWidth: number,
}

const browserHistory = createBrowserHistory({ basename: "" });

/** Component for displaying align goal. */
class AlignGoal extends React.Component<WithTranslation, IState>
{
    localize: TFunction;
    teamId?: string | null = null;
    telemetry?: any = null;
    appInsights: any;
    teamGoalSelectedValue: any;
    teamGoalSelectedId: any;
    personalGoalSelectedValue: any;
    personalGoalSelectedId: any;
    teamGoalCycle: string;

    constructor(props: any) {
        super(props);
        this.localize = this.props.t;

        this.state = {
            loading: false,
            isAlignGoalLoading: false,
            isAlignToAddButtonDisabled: true,
            isAlignGoalButtonDisabled: true,
            isAlignGoalButtonLoading: false,
            isSeeYourGoalButtonDisabled: false,
            isOkayButtonDisabled: false,
            isSeeYourGoalsButtonLoading: false,
            isAlignedGoalsSubmitted: false,
            errorInAddToAlignGoal: "",
            personalGoalSelection: "",
            teamGoalSelection: "",
            teamGoalDetails: [],
            personalGoalDetails: [],
            allPersonalGoalDetails: [],
            alignGoalDetails: [],
            screenWidth: 0,
        };

        let search = window.location.search;
        let params = new URLSearchParams(search);
        this.teamId = params.get("teamId");
        this.telemetry = params.get("telemetry");
        this.teamGoalCycle = "";
    }

    /** Called once component is mounted. */
    async componentDidMount() {
        microsoftTeams.initialize();
        microsoftTeams.getContext(async (context) => {
            this.appInsights = getApplicationInsightsInstance(this.telemetry, browserHistory);

            await this.getTeamGoalDetails();
            await this.getPersonalAndAlignedGoalDetails();
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
    *  Get team goal details from storage.
    * */
    getTeamGoalDetails = async () => {
        this.appInsights.trackTrace({ message: `'getTeamGoalDetails' - Request initiated to fetch team goal details`, severityLevel: SeverityLevel.Information });
        this.setState({ loading: true });
        let teamGoalDetailsResponse = await getTeamGoalDetailsByTeamId(this.teamId);
        if (teamGoalDetailsResponse) {
            if (teamGoalDetailsResponse.status === 200) {
                let teamGoalDetailsData: ITeamGoalProps[] = [];
                teamGoalDetailsResponse.data.forEach((teamGoalDetail) => {
                    teamGoalDetailsData.push({
                        key: teamGoalDetail.TeamGoalId,
                        header: teamGoalDetail.TeamGoalName,
                        TeamGoalId: teamGoalDetail.TeamGoalId,
                        TeamGoalName: teamGoalDetail.TeamGoalName,
                        TeamGoalStartDate: teamGoalDetail.TeamGoalStartDate,
                        TeamGoalEndDate: teamGoalDetail.TeamGoalEndDate,
                    });
                });

                this.setState({ teamGoalDetails: teamGoalDetailsData });

                if (teamGoalDetailsData.length > 0) {
                    // Set team goal cycle to be shown in the align goal task module
                    this.teamGoalCycle = `${moment(teamGoalDetailsData[0].TeamGoalStartDate).format(Constants.goalCycleDateTimeFormat)} ${this.localize("goalCycleToText")} ${moment(teamGoalDetailsData[0].TeamGoalEndDate).format(Constants.goalCycleDateTimeFormat)}`;
                }
            }
            else {
                handleError(teamGoalDetailsResponse);
            }
        }
        this.setState({ loading: false });
    };

    /**
    *  Get team goal details from storage.
    * */
    getPersonalAndAlignedGoalDetails = async () => {
        this.appInsights.trackTrace({ message: `'getPersonalAndAlignedGoalDetails' - Request initiated to fetch personal goal and aligned goal details`, severityLevel: SeverityLevel.Information });
        this.setState({ loading: true });
        let personalGoalDetailsResponse = await getPersonalGoalDetails();
        if (personalGoalDetailsResponse) {
            if (personalGoalDetailsResponse.status === 200) {
                let personalGoals: IPersonalGoalDetail[] = personalGoalDetailsResponse.data;
                personalGoals.forEach((personalGoalDetail) => {
                    personalGoalDetail.TeamGoalName = this.state.teamGoalDetails.find(teamGoalDetail => teamGoalDetail.TeamGoalId === personalGoalDetail.TeamGoalId)?.TeamGoalName
                });
                this.setState({ allPersonalGoalDetails: personalGoals });

                if (personalGoals.length > 0) {
                    this.appInsights.trackTrace({ message: `'getPersonalAndAlignedGoalDetails' - Request initiated to add align goals`, severityLevel: SeverityLevel.Information });
                    let alignedGoalDetails: IPersonalGoalDetail[] = [];
                    personalGoals.forEach((personalGoalDetail) => {
                        if (personalGoalDetail.IsAligned) {
                            personalGoalDetail.TeamGoalId?.split(",").forEach((alignedTeamGoalId) => {
                                let alignedGoal: IPersonalGoalDetail = {} as IPersonalGoalDetail;
                                alignedGoal.UserAadObjectId = personalGoalDetail.UserAadObjectId;
                                alignedGoal.AdaptiveCardActivityId = personalGoalDetail.AdaptiveCardActivityId;
                                alignedGoal.ConversationId = personalGoalDetail.ConversationId;
                                alignedGoal.CreatedOn = personalGoalDetail.CreatedOn;
                                alignedGoal.CreatedBy = personalGoalDetail.CreatedBy;
                                alignedGoal.LastModifiedOn = personalGoalDetail.LastModifiedOn;
                                alignedGoal.LastModifiedBy = personalGoalDetail.LastModifiedBy;
                                alignedGoal.IsActive = personalGoalDetail.IsActive;
                                alignedGoal.IsAligned = personalGoalDetail.IsAligned;
                                alignedGoal.IsDeleted = personalGoalDetail.IsDeleted;
                                alignedGoal.IsReminderActive = personalGoalDetail.IsReminderActive;
                                alignedGoal.GoalName = personalGoalDetail.GoalName;
                                alignedGoal.PersonalGoalId = personalGoalDetail.PersonalGoalId;
                                alignedGoal.ReminderFrequency = personalGoalDetail.ReminderFrequency;
                                alignedGoal.Status = personalGoalDetail.Status;
                                alignedGoal.StartDate = personalGoalDetail.StartDate;
                                alignedGoal.EndDate = personalGoalDetail.EndDate;
                                alignedGoal.ServiceURL = personalGoalDetail.ServiceURL;
                                alignedGoal.TeamId = personalGoalDetail.TeamId;
                                alignedGoal.TeamGoalId = alignedTeamGoalId;
                                alignedGoal.TeamGoalName = this.state.teamGoalDetails.find(teamGoalDetail => teamGoalDetail.TeamGoalId === alignedTeamGoalId)?.TeamGoalName;
                                alignedGoal.EndDateUTC = personalGoalDetail.EndDateUTC;
                                alignedGoal.NotesCount = personalGoalDetail.NotesCount;
                                alignedGoal.GoalCycleId = personalGoalDetail.GoalCycleId;

                                alignedGoalDetails.push(alignedGoal);
                            });
                        }
                    });
                    this.setState({ alignGoalDetails: alignedGoalDetails });
                }

                let personalGoalsDataForDropDown: IPersonalGoalProps[] = [];
                personalGoals.forEach((personalGoalDetail) => {
                    personalGoalsDataForDropDown.push({
                        key: personalGoalDetail.PersonalGoalId,
                        header: personalGoalDetail.GoalName,
                        PersonalGoalId: personalGoalDetail.PersonalGoalId,
                        GoalName: personalGoalDetail.GoalName,
                    });
                });
                this.setState({ personalGoalDetails: personalGoalsDataForDropDown });
            }
            else {
                handleError(personalGoalDetailsResponse);
            }
        }
        this.setState({ loading: false });
    };

    /**
    *  Gets called when user select a personal goal.
    * */
    private onPersonalGoalSelected = {
        onAdd: personalGoal => {
            this.personalGoalSelectedValue = personalGoal.GoalName;
            this.personalGoalSelectedId = personalGoal.PersonalGoalId;
            this.setState({ personalGoalSelection: this.personalGoalSelectedValue });

            if (this.personalGoalSelectedValue && this.state.teamGoalSelection) {
                if (!this.state.errorInAddToAlignGoal) {
                    this.setState({ isAlignToAddButtonDisabled: false });
                }
            }
            return "";
        }
    }

    /**
    *  Gets called when user select a team goal.
    * */
    private onTeamGoalSelected = {
        onAdd: teamGoal => {
            this.teamGoalSelectedValue = teamGoal.TeamGoalName;
            this.teamGoalSelectedId = teamGoal.TeamGoalId;
            this.setState({ teamGoalSelection: this.teamGoalSelectedValue });
            this.setState({ errorInAddToAlignGoal: "" });

            if (this.state.personalGoalSelection && this.teamGoalSelectedValue) {
                this.setState({ isAlignToAddButtonDisabled: false });
            }
            return "";
        }
    }

    /**
    *  Gets called when user clicks on add to align goal button.
    * */
    private addToAlignGoalsButtonClick = async () => {
        this.appInsights.trackTrace({ message: `'addToAlignGoalsButtonClick' - Request initiated add to align goal button is clicked`, severityLevel: SeverityLevel.Information });
        let alignedGoalDetail: IPersonalGoalDetail = {} as IPersonalGoalDetail;
        let teamGoalAlreadyAlign: IPersonalGoalDetail = {} as IPersonalGoalDetail;
        alignedGoalDetail = this.state.allPersonalGoalDetails.find(alignedGoalDetail => alignedGoalDetail.PersonalGoalId === this.personalGoalSelectedId) as IPersonalGoalDetail;
        teamGoalAlreadyAlign = this.state.alignGoalDetails.find(teamGoalDetail => teamGoalDetail.TeamGoalId === this.teamGoalSelectedId) as IPersonalGoalDetail;
        if (teamGoalAlreadyAlign) {
            this.setState({ errorInAddToAlignGoal: this.localize("alignGoalErrorOneTeamGoalPerPersonalGoalLimitation"), isAlignToAddButtonDisabled: true });
        }
        else {
            let newAlignGoalDetail: IPersonalGoalDetail =
            {
                UserAadObjectId: alignedGoalDetail.UserAadObjectId,
                AdaptiveCardActivityId: alignedGoalDetail.AdaptiveCardActivityId,
                ConversationId: alignedGoalDetail.ConversationId,
                IsActive: alignedGoalDetail.IsActive,
                IsAligned: true,
                IsDeleted: alignedGoalDetail.IsDeleted,
                IsReminderActive: alignedGoalDetail.IsReminderActive,
                CreatedOn: alignedGoalDetail.CreatedOn,
                CreatedBy: alignedGoalDetail.CreatedBy,
                LastModifiedOn: alignedGoalDetail.LastModifiedOn,
                LastModifiedBy: alignedGoalDetail.LastModifiedBy,
                GoalName: this.personalGoalSelectedValue,
                PersonalGoalId: alignedGoalDetail.PersonalGoalId,
                ReminderFrequency: alignedGoalDetail.ReminderFrequency,
                Status: alignedGoalDetail.Status,
                StartDate: alignedGoalDetail.StartDate,
                EndDate: alignedGoalDetail.EndDate,
                ServiceURL: alignedGoalDetail.ServiceURL,
                TeamId: this.teamId,
                TeamGoalId: this.state.teamGoalDetails.find(teamGoalDetail => teamGoalDetail.TeamGoalId === this.teamGoalSelectedId)?.TeamGoalId,
                TeamGoalName: this.teamGoalSelectedValue,
                EndDateUTC: alignedGoalDetail.EndDateUTC,
                NotesCount: alignedGoalDetail.NotesCount,
                GoalCycleId: alignedGoalDetail.GoalCycleId,
            };
            this.state.alignGoalDetails.push(newAlignGoalDetail);
            this.setState({ errorInAddToAlignGoal: "", personalGoalSelection: "", teamGoalSelection: "" });
            this.setState({ isAlignToAddButtonDisabled: true, isAlignGoalButtonDisabled: false });
        }
        this.setState({ loading: false });
    }

    /**
    *  Gets called when user clicks on close button in front of aligned goals.
    * */
    private removeFromAlignedGoal = (teamGoalId?: string | null) => () => {
        if (this.state.alignGoalDetails) {
            this.appInsights.trackTrace({ message: `'removeFromAlignedGoal' - Request initiated cross button clicked to remove align goal detail`, severityLevel: SeverityLevel.Information });
            let updatedAlignedGoalDetails = this.state.alignGoalDetails.filter(alignedGoalDetail => alignedGoalDetail.TeamGoalId !== teamGoalId);
            this.setState({ alignGoalDetails: updatedAlignedGoalDetails });
            this.setState({ isAlignGoalButtonDisabled: false, errorInAddToAlignGoal: "" })
            if (this.state.personalGoalSelection && this.state.teamGoalSelection) {
                this.setState({ isAlignToAddButtonDisabled: false });
            }
        }

        this.setState({ loading: false });
    }

    /**
    *  Gets called when user clicks on align goals button.
    * */
    private onAlignGoalsButtonClick = async () => {
        this.appInsights.trackTrace({ message: `'onAlignGoalsButtonClick' - Request initiated align goals button is clicked`, severityLevel: SeverityLevel.Information });
        this.setState({ isAlignToAddButtonDisabled: true, isAlignGoalButtonDisabled: true, isAlignGoalButtonLoading: true });

        let response = await this.saveGoalAlignmentDetails();
        if (response) {
            this.setState({ isAlignGoalButtonLoading: false, isAlignedGoalsSubmitted: true });
        }
        else {
            this.setState({ errorInAddToAlignGoal: this.localize("alignGoalErrorInSavingAligedGoalDetails"), isAlignGoalButtonDisabled: false });
        }

        this.setState({ errorInAddToAlignGoal: "", personalGoalSelection: "", teamGoalSelection: "" });
        this.setState({ isOkayButtonDisabled: false, isSeeYourGoalButtonDisabled: false, isSeeYourGoalsButtonLoading: false });
        this.setState({ loading: false });
    }

    /**
    *  Stores personal and aligned goal details in table storage.
    * */
    private saveGoalAlignmentDetails = async () => {
        if (this.state.allPersonalGoalDetails.length > 0) {
            this.appInsights.trackTrace({ message: `'saveGoalAlignmentDetails' - Request initiated saving aligned and unaligned goal details in storage`, severityLevel: SeverityLevel.Information });
            let personalGoalDetailsData: IPersonalGoalDetail[] = [];
            let personalGoalDetails = this.state.allPersonalGoalDetails;
            personalGoalDetails.forEach((personalGoalDetail) => {
                let alignedGoalDetail = this.state.alignGoalDetails.find(alignedGoalDetail => alignedGoalDetail.PersonalGoalId === personalGoalDetail.PersonalGoalId);
                if (alignedGoalDetail) {
                    let commaSeparatedAlignedTeamGoalId: string = "";
                    let alignedGoalDetails = this.state.alignGoalDetails.filter(alignedGoal => alignedGoal.PersonalGoalId === personalGoalDetail.PersonalGoalId);
                    alignedGoalDetails.forEach((alignedGoal) => {
                        commaSeparatedAlignedTeamGoalId += alignedGoal.TeamGoalId + ",";
                    });
                    personalGoalDetail = alignedGoalDetail;
                    personalGoalDetail.LastModifiedOn = new Date().toUTCString();
                    personalGoalDetail.TeamGoalId = commaSeparatedAlignedTeamGoalId.slice(0, -1);
                    personalGoalDetailsData.push(personalGoalDetail);
                }
                else {
                    personalGoalDetail.IsAligned = false;
                    personalGoalDetail.LastModifiedOn = new Date().toUTCString();
                    personalGoalDetail.TeamId = null;
                    personalGoalDetail.TeamGoalId = null;
                    personalGoalDetail.TeamGoalName = null;
                    personalGoalDetailsData.push(personalGoalDetail);
                }
            });
            const savePersonalGoalDetailsResponse = await savePersonalGoalDetails(personalGoalDetailsData)
            if (savePersonalGoalDetailsResponse.status !== 200 && savePersonalGoalDetailsResponse.status !== 204) {
                this.setState({ isAlignGoalButtonLoading: false, errorInAddToAlignGoal: this.localize("alignGoalErrorInSavingAligedGoalDetails") });
                handleError(savePersonalGoalDetailsResponse);
                return false;
            }
        }

        this.setState({ isAlignGoalButtonLoading: true });
        this.appInsights.trackTrace({ message: `'saveGoalAlignmentDetails' - Personal goals aligned successfully and teamId=${this.teamId}`, severityLevel: SeverityLevel.Information });
        return true;
    }

    /**
    *  Gets called when user clicks on okay button in aligned goal success page.
    * */
    private onOkayButtonClick = () => {
        this.appInsights.trackTrace({ message: `'onOkayButtonClick' - Request initiated okay button is clicked`, severityLevel: SeverityLevel.Information });
        this.setState({ isSeeYourGoalButtonDisabled: true });

        microsoftTeams.getContext((context) => {
            microsoftTeams.tasks.submitTask();
        });

        this.setState({ loading: false });
    }

    /**
    *  Gets called when user clicks on see your goals button on align goal success screen.
    * */
    private onSeeYourGoalsButtonClick = () => {
        this.appInsights.trackTrace({ message: `'onSeeYourGoalsButtonClick' - Request initiated see your goals button is clicked`, severityLevel: SeverityLevel.Information });
        this.setState({ isOkayButtonDisabled: true, isSeeYourGoalsButtonLoading: true, isAlignedGoalsSubmitted: false });
        this.setState({ loading: false });
    }

    /**
    *  Renders align goal layout on UI.
    * */
    renderAlignGoal() {
        return (
            <>
                <Flex gap="gap.large" vAlign="center" className="align-goal-cycle-heading">
                    {this.teamGoalCycle && <Text weight="bold" align="center" content={`${this.localize("goalCycleText")}: ${this.teamGoalCycle}`} />}
                </Flex>
                <Flex gap="gap.large" vAlign="center" className="align-goal-team-dropdown-title">
                    <Text content={this.localize("alignGoalTeamGoalTitleText")} />
                </Flex>
                <div className={this.state.screenWidth <= 599 ? "error-in-align-goal-for-small-device" : "error-in-align-goal"}>
                    <Flex gap="gap.small" hAlign="end">
                        {this.state.errorInAddToAlignGoal !== null && <Text content={this.state.errorInAddToAlignGoal} error />}
                    </Flex>
                </div>
                <div>
                    <Dropdown
                        placeholder={this.localize("alignGoalTeamGoalDropDownPlaceholderText")}
                        items={this.state.teamGoalDetails}
                        value={this.state.teamGoalSelection}
                        noResultsMessage={this.localize("alignGoalTeamGoalDropDownNoResultText")}
                        getA11ySelectionMessage={this.onTeamGoalSelected}
                        disabled={this.state.isAlignGoalButtonLoading}
                        fluid
                        checkable
                    />
                </div>
                <Flex gap="gap.large" vAlign="center" className="align-goal-personal-dropdown-title">
                    <Text content={this.localize("alignGoalPersonalGoalTitleText")} />
                </Flex>
                <div>
                    <Dropdown
                        placeholder={this.localize("alignGoalPersonalGoalDropDownPlaceholderText")}
                        items={this.state.personalGoalDetails}
                        value={this.state.personalGoalSelection}
                        noResultsMessage={this.localize("alignGoalPersonalGoalDropDownPlaceholderText")}
                        getA11ySelectionMessage={this.onPersonalGoalSelected}
                        disabled={this.state.isAlignGoalButtonLoading}
                        fluid
                        checkable
                    />
                </div>
                <div className="add-to-align-button-middle" >
                    <Flex>
                        <Flex.Item align="end" size="size.small" >
                            <Button content={this.localize("alignGoalAddToAlignButtonText")} primary onClick={this.addToAlignGoalsButtonClick} disabled={this.state.isAlignToAddButtonDisabled} />
                        </Flex.Item>
                    </Flex>
                </div>
                <div className="align-goal-table-bottom-content" >
                    <Divider />
                    {this.state.alignGoalDetails.length > 0 && this.state.screenWidth <= 599 &&
                        <div>
                            <Table aria-label="table" className="aligned-goal-table-for-small-device">
                                {this.state.alignGoalDetails.map((alignedGoal) => (
                                    <div>
                                        <Flex gap="gap.smaller" vAlign="center">
                                            <div className="align-goal-overflow-personalgoal-for-small-device">
                                                <Flex.Item align="start" size="size.small" grow>
                                                    <Text content={alignedGoal.TeamGoalName} title={alignedGoal.TeamGoalName!} />
                                                </Flex.Item>
                                            </div>
                                            <Flex.Item align="end">
                                                <Separator vertical />
                                            </Flex.Item>
                                            <div>
                                                <Flex.Item align="end">
                                                    <Button circular icon={<TrashCanIcon />} text iconOnly title={this.localize("alignGoalCloseIconToottipText")} onClick={this.removeFromAlignedGoal(alignedGoal.TeamGoalId)} disabled={this.state.isAlignGoalButtonLoading} />
                                                </Flex.Item>
                                            </div>
                                        </Flex>
                                        <Flex gap="gap.smaller" vAlign="center">
                                            <div className="align-goal-overflow-teamgoal-for-small-device">
                                                <Flex.Item align="start" size="size.small" grow>
                                                    <Text content={alignedGoal.GoalName} title={alignedGoal.GoalName} />
                                                </Flex.Item>
                                            </div>
                                        </Flex>
                                        <Separator />
                                    </div>
                                ))}
                            </Table>
                        </div>
                    }
                    {this.state.alignGoalDetails.length === 0 && this.state.screenWidth <= 559 &&
                        <Flex gap="gap.large" hAlign="center" vAlign="center" className="align-goal-bottom-content-no-goals-for-small-device" >
                            <Flex.Item align="center" size="size.small" grow>
                                <Text content={this.localize("alignGoalEmptyTableContent")} />
                            </Flex.Item>
                        </Flex>
                    }
                    {this.state.alignGoalDetails.length > 0 && this.state.screenWidth > 599 &&
                        <div>
                            <div className="align-goal-table-bottom-content-header" >
                                <Flex gap="gap.large" vAlign="center">
                                    <Flex.Item align="start" size="size.small" grow>
                                        <Text content={this.localize("alignGoalTableTitleTeamGoalText")} />
                                    </Flex.Item>
                                    <Flex.Item align="start" size="size.small" grow>
                                        <Text content={this.localize("alignGoalTableTitlePersonalGoalText")} />
                                    </Flex.Item>
                                    <Flex.Item align="end" className="align-goal-close-button-margin" >
                                        <Button text iconOnly title="Close" />
                                    </Flex.Item>
                                </Flex>
                            </div>
                            <Table aria-label="table" className="aligned-goal-table">
                                {this.state.alignGoalDetails.map((alignedGoal) => (
                                    <div>
                                        <Flex gap="gap.large" vAlign="center">
                                            <div className="align-goal-table-overflow-text-left-side">
                                                <Flex.Item align="start" size="size.small" grow>
                                                    <Text content={alignedGoal.TeamGoalName} title={alignedGoal.TeamGoalName!} />
                                                </Flex.Item>
                                            </div>
                                            <div className="align-goal-table-overflow-text-right-side">
                                                <Flex.Item align="start" size="size.small" grow>
                                                    <Text content={alignedGoal.GoalName} title={alignedGoal.GoalName} />
                                                </Flex.Item>
                                            </div>
                                            <Flex.Item align="end" className="align-goal-close-button-margin" >
                                                <Button circular icon={<CloseIcon />} text iconOnly title={this.localize("alignGoalCloseIconToottipText")} onClick={this.removeFromAlignedGoal(alignedGoal.TeamGoalId)} disabled={this.state.isAlignGoalButtonLoading} />
                                            </Flex.Item>
                                        </Flex>
                                    </div>
                                ))}
                            </Table>
                        </div>
                    }
                    {this.state.alignGoalDetails.length === 0 && this.state.screenWidth > 599 &&
                        <Flex gap="gap.large" hAlign="center" vAlign="center" className="align-goal-bottom-content-no-goals" >
                            <Flex.Item align="center" size="size.small" grow>
                                <Text content={this.localize("alignGoalEmptyTableContent")} />
                            </Flex.Item>
                        </Flex>
                    }
                    <div className="align-goals-button-bottom">
                        <Flex.Item align="end" size="size.small" >
                            <Button content={this.localize("alignGoalAlignGoalsButtonText")} primary onClick={this.onAlignGoalsButtonClick} disabled={this.state.isAlignGoalButtonDisabled} loading={this.state.isAlignGoalButtonLoading} />
                        </Flex.Item>
                    </div>
                </div>
            </>
        );
    }

    /**
    *  Renders align goal layout or loader on UI depending upon data is fetched from storage.
    * */
    render() {
        let contents = this.state.isAlignedGoalsSubmitted
            ? <AlignGoalSuccessScreen
                onSeeYourGoalsButtonClick={this.onSeeYourGoalsButtonClick}
                onOkayButtonClick={this.onOkayButtonClick}
                isSeeYourGoalButtonDisabled={this.state.isSeeYourGoalButtonDisabled}
                isSeeYourGoalsButtonLoading={this.state.isSeeYourGoalsButtonLoading}
                isOkayButtonDisabled={this.state.isOkayButtonDisabled}
            />
            : this.renderAlignGoal();
        if (!this.state.loading) {
            return (
                <div className="container-div">
                    {contents}
                </div>
            );
        }
        else {
            return (
                <Loader />
            );
        }
    }
}

export default withTranslation()(AlignGoal);