// <copyright file="personal-goal-notes.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import * as React from "react";
import { Text, Button, Flex, Input } from "@fluentui/react-northstar";
import { TrashCanIcon, EditIcon } from '@fluentui/react-icons-northstar';
import { useTranslation } from 'react-i18next';
import moment from "moment";
import { IPersonalGoalNoteDetail } from "../../models/type";
import Constants from "../../constants";
import "../../styles/style.css";

interface IPersonalGoalNoteProps {
    notesData: IPersonalGoalNoteDetail[],
    onGoalNoteDeleteButtonClick: (goalNoteId: string) => void,
    onGoalNoteEditButtonClick: (index: number, event: any) => void
    handleGoalNoteChange: (index: number, event: any) => void
}

const PersonalGoalNote: React.FunctionComponent<IPersonalGoalNoteProps> = props => {
    const { t } = useTranslation();
    return (
        <>
            {props.notesData.length > 0 &&
                <Flex gap="gap.large" vAlign="center" className="control-spacing">
                    <Text weight="bold" align="center" className="note-header" content={t("goalNotesText") + " (" + props.notesData.length + ")"} />
                </Flex> 
            }
            {props.notesData.map((note, id) => (
                <div>
                    <Flex space="between">
                        <Text className="note-table-header control-spacing truncate-source-name" weight="semilight" align="start" content={note.SourceName} title={note.SourceName} />
                        <Flex gap="gap.small">
                            <Text className="note-table-header control-spacing" weight="light" align="end" content={moment(note.CreatedOn).format('LLL')} />
                            <Button size="smaller" className="padding-small" icon={<EditIcon />} text iconOnly title={t("editGoalText")} id={id + ""} onClick={event => props.onGoalNoteEditButtonClick(id, event)} />
                            <Button size="smaller" className="padding-small" icon={<TrashCanIcon />} text iconOnly title={t("deleteGoalText")} onClick={event => props.onGoalNoteDeleteButtonClick(note.PersonalGoalNoteId)} />
                        </Flex>
                    </Flex>
                    <Flex gap="gap.large" vAlign="start" className="control-padding">
                        {note.IsEdited ? <Input maxLength={Constants.maxAllowedNoteDescription} fluid className="add-goals-input" value={note.PersonalGoalNoteDescription} title={note.PersonalGoalNoteDescription} onChange={event => props.handleGoalNoteChange(id, event)} />
                            : <Text align="start" content={note.PersonalGoalNoteDescription} className="note-text" />}
                    </Flex>
                </div>
             ))}
        </>
    );
}

export default PersonalGoalNote;