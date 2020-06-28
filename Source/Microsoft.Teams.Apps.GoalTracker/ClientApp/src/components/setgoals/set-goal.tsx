// <copyright file="set-goal.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import React from 'react';
import { Text, Button, Flex, Checkbox, RadioGroup, Input, List } from '@fluentui/react-northstar';
import "../../styles/style.css";
import { AddIcon } from '@fluentui/react-icons-northstar';
import StartDateEndDate from './date-picker';
import { useTranslation } from 'react-i18next';
import Constants from "../../constants";
import { getReminderFrequncyCollection } from "../../helpers/goal-helper";
interface ISetGoalsProps {
    errorMessage: string,
    goals: any,
    isReminderActive: boolean
    getStartDate: (startDate: Date | undefined) => void,
    getEndDate: (endDate: Date | undefined) => void,
    setIsReminderActive: (event: any, checkboxProps: any) => void,
    setReminder: (event: any, props) => void,
    saveGoals: (event: any) => void,
    removeGoals: (event: any) => void,
    addNewGoal: (event: any) => void,
    addGoalFromTextBox: (event: any) => void,
    goalName: string,
    startDate: string,
    minStartDate: string,
    endDate: string,
    reminderFrequency: number,
    isSaveButtonLoading: boolean,
    isSaveButtonDisabled: boolean,
    theme?: string | null,
    screenWidth: number,
}
const SetGoal: React.FunctionComponent<ISetGoalsProps> = props => {
    const { t } = useTranslation();
    return (
        <div>
            <div className="padding-small">{t('addGoalTaskmoduleText')}</div>
            <div className="padding-small">
                <Flex gap="gap.smaller">
                    <Text content={t('goalNameLabel')} />
                </Flex>
                {props.screenWidth <= 599 &&
                    <div>
                        <Flex className="goal-list-for-small-device" >
                            <List truncateHeader={true} items={props.goals} />
                        </Flex>
                        <div className="add-goal">
                            <Input fluid className="add-goals-input-main-for-small-device" placeholder={t('addGoalPlaceHolder')} value={props.goalName} onChange={event => props.addGoalFromTextBox(event)} maxLength={Constants.maxAllowedGoalName} title={props.goalName} />
                            <Button content={t('addButtonText')} className="add-goal-button-for-small-device" aria-label={t('addGoalIcon')} onClick={props.addNewGoal} secondary />
                        </div>
                        <div className="padding-small-for-small-device">
                            <StartDateEndDate
                                getStartDate={props.getStartDate}
                                getEndDate={props.getEndDate}
                                startDate={props.startDate}
                                minStartDate={props.minStartDate}
                                endDate={props.endDate}
                                theme={props.theme!}
                                screenWidth={props.screenWidth}
                            />
                        </div>
                        <div className="padding-small"> <Checkbox label={t('remindMeLabel')} checked={props.isReminderActive} onClick={props.setIsReminderActive} /></div>
                        <div className="radio-group-spacing-for-small-device">
                            <RadioGroup
                                defaultCheckedValue={1}
                                items={getReminderFrequncyCollection(t, !props.isReminderActive)}
                                onCheckedValueChange={props.setReminder}
                                checkedValue={props.reminderFrequency}
                            />
                        </div>
                    </div>
                }
                {props.screenWidth > 599 &&
                    <div>
                        <Flex className="goal-list" >
                            <List truncateHeader={true} items={props.goals} />
                        </Flex>
                        <div className="add-goal">
                            <Input fluid className="add-goals-input-main" placeholder={t('addGoalPlaceHolder')} value={props.goalName} onChange={event => props.addGoalFromTextBox(event)} maxLength={Constants.maxAllowedGoalName} title={props.goalName} />
                            <Button content={t('addButtonText')} className="add-goal-button" aria-label={t('addGoalIcon')} onClick={props.addNewGoal} secondary />
                        </div>
                        <div className="padding-small">
                            <StartDateEndDate
                                getStartDate={props.getStartDate}
                                getEndDate={props.getEndDate}
                                startDate={props.startDate}
                                minStartDate={props.minStartDate}
                                endDate={props.endDate}
                                theme={props.theme!}
                                screenWidth={props.screenWidth}
                            />
                        </div>
                        <div className="padding-small"> <Checkbox label={t('remindMeLabel')} checked={props.isReminderActive} onClick={props.setIsReminderActive} /></div>
                        <div className="radio-group-spacing">
                            <RadioGroup
                                defaultCheckedValue={1}
                                items={getReminderFrequncyCollection(t, !props.isReminderActive)}
                                onCheckedValueChange={props.setReminder}
                                checkedValue={props.reminderFrequency}
                            />
                        </div>
                    </div>
                }
            </div>
            <div className="footer">
                <div className="error">
                    <Flex gap="gap.small">
                        {props.errorMessage !== null && <Text className="small-margin-left" content={props.errorMessage} error />}
                    </Flex>
                </div>
                <div className="save-button">
                    <Flex gap="gap.smaller">
                        <Button content={t('saveButtonText')} primary loading={props.isSaveButtonLoading} disabled={props.isSaveButtonDisabled} onClick={props.saveGoals} />
                    </Flex>
                </div>
            </div>
        </div>
    );
}
export default SetGoal;