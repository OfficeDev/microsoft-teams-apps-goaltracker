// <copyright file="align-goal-success-screen.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import React from 'react';
import { Text, Button, Flex, Image } from '@fluentui/react-northstar';
import "../../styles/style.css";
import { useTranslation } from 'react-i18next';


interface IState {
    onSeeYourGoalsButtonClick: (event: any, props) => void
    onOkayButtonClick: (event: any, props) => void
    isSeeYourGoalButtonDisabled: boolean,
    isSeeYourGoalsButtonLoading: boolean,
    isOkayButtonDisabled: boolean
}

const AlignGoalSuccessScreen: React.FunctionComponent<IState> = props => {
    const { t } = useTranslation();
    return (
        <>
            <div className="align-goal-success-message">
                <div>
                    <Flex gap="gap.large" vAlign="center" hAlign="center">
                        <Image avatar src="/Artifacts/alignGoalSuccessIcon.png" />
                    </Flex>
                </div>
                <div>
                    <Flex gap="gap.large" vAlign="center" hAlign="center">
                        <Text content={t('alignGoalWellDoneText')} />
                    </Flex>
                </div>
                <div>
                    <Flex gap="gap.large" vAlign="center" hAlign="center">
                        <Text content={t('alignGoalSuccessMessageText')} />
                    </Flex>
                </div>
                <div className="align-goal-back-button-bottom">
                    <Flex.Item align="end" size="size.small" >
                        <Button content={t('alignGoalSeeYourGoalButtonText')} secondary onClick={props.onSeeYourGoalsButtonClick} disabled={props.isSeeYourGoalButtonDisabled} loading={props.isSeeYourGoalsButtonLoading} />
                    </Flex.Item>
                </div>
            </div>
            <div className="align-goals-button-bottom">
                <Flex.Item align="end" size="size.small" >
                    <Button content={t('alignGoalOkayButtonText')} primary onClick={props.onOkayButtonClick} disabled={props.isOkayButtonDisabled} />
                </Flex.Item>
            </div>
        </>
    );
}

export default AlignGoalSuccessScreen;