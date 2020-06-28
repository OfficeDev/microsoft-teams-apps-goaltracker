// <copyright file="personal-goal-table.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import * as React from "react";
import { Table, Text, Button, Accordion, Dialog, Flex, Label, List, Divider } from "@fluentui/react-northstar";
import { TrashCanIcon } from '@fluentui/react-icons-northstar';
import { useTranslation } from 'react-i18next';
import { IPersonalGoalDetail } from "../../models/type";
import "../../styles/style.css";
import { Separator } from "office-ui-fabric-react";

interface IPersonalGoalsTableProps {
    goalsData: IPersonalGoalDetail[],
    goalStatus: any,
    onDeleteButtonClick: (goalDetails: IPersonalGoalDetail) => boolean,
    onPersonalGoalClick: (goalId: string, t: any) => void,
    screenWidth: number,
}

const PersonalGoalTable: React.FunctionComponent<IPersonalGoalsTableProps> = props => {
    const { t } = useTranslation();
    const goalTableHeader = {
        key: "header",
        items:
            [
                { content: <Text weight="regular" content={t('goalNameHeader')} />, className: "table-header goal-table-goal-name" },
                { content: <Text weight="regular" content={t('alignedWithHeader')} />, className: "table-header goal-table-align-with" },
                { content: <Text weight="regular" content={t('goalNoteCountHeader')} />, className: "table-header goal-table-note" },
                { content: <Text weight="regular" content="" />, className: "goal-table-delete" }
            ]
    };

    let goalDataRowsForDesktop = props.goalsData.map((value: any, index) => (
        {
            key: value.GoalId,
            GoalStatus: value.Status,
            style: {},
            items:
                [
                    { content: <Text weight="semibold" content={value.GoalName} title={value.GoalName} />, key: index + "2", truncateContent: true, className: "table-row goal-table-goal-name-cell", onClick:() => props.onPersonalGoalClick(value.PersonalGoalId, t) },
                    { content: <Text content={value.TeamGoalName} title={value.TeamGoalName} />, key: index + "3", truncateContent: true, className: "table-row goal-table-align-with-cell", onClick: () => props.onPersonalGoalClick(value.PersonalGoalId, t) },
                    { content: <Text content={value.NotesCount} title={value.NotesCount} />, key: index + "4", truncateContent: true, className: "table-row goal-table-note-cell", onClick: () => props.onPersonalGoalClick(value.PersonalGoalId, t) },
                    {
                        content:
                            <Dialog
                                cancelButton={t('cancelButtonText')}
                                confirmButton={t('confirmButtonText')}
                                onConfirm={() => props.onDeleteButtonClick(value)}
                                content={t('goalDeleteConfirmationMessageText')}
                                header={t('actionConfirmationMessage')}
                                trigger={<Button size="smaller" text iconOnly icon={<TrashCanIcon />} title={t('deleteGoalText')} />}
                                className="goal-delete-confirmation-dialog"
                            />, className: "goal-table-delete"
                    }
                ]
        }
    ));

    let goalDataRowsListView = props.goalsData.map((value: any) => (
        {
            key: value.GoalId,
            GoalStatus: value.Status,
            header: <></>,
            content:
                <>
                    <Flex vAlign="stretch">
                        <div className="goal-list-for-small-device">
                            <Flex.Item>
                                <Flex column gap="gap.small" vAlign="stretch">
                                    <Flex>
                                        <Text className="goal-heading" onClick={() => props.onPersonalGoalClick(value.PersonalGoalId, t)} title={value.GoalName} content={value.GoalName} />
                                    </Flex>
                                    <div className="aligned-unaligned-text">
                                        <Flex vAlign="start">
                                            <Label content={value.TeamGoalName === t('notAlignedTeamGoaltext') ? t('notAlignedTeamGoaltext') : t('alignedTeamGoaltext')} title={value.TeamGoalName} circular />
                                            <Text content={value.TeamGoalName === t('notAlignedTeamGoaltext') ? "gap" : "glarge"} className="note-gap-text" />
                                            <Text content={`${value.NotesCount} (${t('goalNotesText')})`} title={value.NotesCount} />
                                        </Flex>
                                    </div>
                                </Flex>
                            </Flex.Item>
                        </div>
                        <Separator vertical />
                        <div className="delete-icon-for-small-device">
                            <Flex.Item push align="end">
                                <Dialog
                                    className="dialog-container-goal-list"
                                    cancelButton={t('cancelButtonText')}
                                    confirmButton={t('confirmButtonText')}
                                    content={t('goalDeleteConfirmationMessageText')}
                                    header={t('actionConfirmationMessage')}
                                    trigger={<Button size="smaller" text iconOnly icon={<TrashCanIcon />} title={t('deleteGoalText')} />}
                                    onConfirm={() => props.onDeleteButtonClick(value)}
                                />
                            </Flex.Item>
                        </div>
                    </Flex>
                    <div className="goal-list-divider-for-small-device">
                        <Divider />
                    </div>
                </>
        }
    ));

    let panelsForDesktop = props.goalStatus.map((goalStatus: any) => (
        {
            title: <Text content={goalStatus.DisplayName + " (" + goalDataRowsForDesktop.filter(row => row.GoalStatus === goalStatus.value).length + ")"} className="goal-header" />,
            content: <Table rows={goalDataRowsForDesktop.filter(row => row.GoalStatus === goalStatus.value)} header={goalTableHeader} className="table-cell-content" />
        }
    ));

    let panelsForListItem = props.goalStatus.map((goalStatus: any) => (
        {
            title: <Text content={goalStatus.DisplayName + " (" + goalDataRowsListView.filter(row => row.GoalStatus === goalStatus.value).length + ")"} className="goal-header-for-small-device" />,
            content: <List items={goalDataRowsListView.filter(row => row.GoalStatus === goalStatus.value)} className="table-cell-content-for-small-device" />
        }
    ));

    return (
        <>
            {props.screenWidth <= 750 && <Accordion defaultActiveIndex={[0]} panels={panelsForListItem} className="accordian" />}
            {props.screenWidth > 750 && <Accordion defaultActiveIndex={[0]} panels={panelsForDesktop} className="accordian" />}
        </>
    );
}

export default PersonalGoalTable;