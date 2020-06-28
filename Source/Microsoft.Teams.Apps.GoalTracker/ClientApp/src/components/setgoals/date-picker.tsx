// <copyright file="date-picker.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import React from 'react';
import moment from "moment";
import { Flex } from '@fluentui/react-northstar';
import { useState } from "react";
import { useTranslation } from "react-i18next";
import { DatePicker } from 'office-ui-fabric-react/lib/DatePicker';
import { Fabric, Customizer } from 'office-ui-fabric-react/lib';
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import { DarkCustomizations } from "../../helpers/theme/dark-customizations";
import { DefaultCustomizations } from "../../helpers/theme/default-customizations";
import Constants from "../../constants";
import "../../styles/style.css";

initializeIcons();

interface IDateePickerProps {
    getStartDate: (startDate: Date | undefined) => void,
    getEndDate: (endDate: Date | undefined) => void
    startDate: string,
    minStartDate: string,
    endDate: string,
    theme: string
    screenWidth: number,
}

const StartDateEndDate: React.FC<IDateePickerProps> = props => {
    let theme = props.theme;
    let datePickerTheme;
    if (theme === Constants.dark) {
        datePickerTheme = DarkCustomizations
    }
    else if (theme === Constants.contrast) {
        datePickerTheme = DarkCustomizations
    }
    else {
        datePickerTheme = DefaultCustomizations
    }
    const { t } = useTranslation();
    const [minEndDate, setMinEndDate] = useState<Date>(new Date(moment().add(30, 'd').format()));

    /**
    * Handle change event for goal cycle start date picker.
    * @param date | cycle start date.
    */
    const onSelectStartDate = (date: Date | null | undefined): void => {
        if (date) {
            let startCycle = moment(date)
                .set('hour', moment().hour())
                .set('minute', moment().minute())
                .set('second', moment().second())
            //setStartDate(date);
            setMinEndDate(new Date(moment(startCycle.toDate()).add(30, 'd').format()));
            props.getStartDate(startCycle.toDate()!);
        }
        else {
            props.getStartDate(undefined);
        }
    };

    /**
     * Handle change event for goal cycle end date picker.
     * @param date | cycle end date.
     */
    const onSelectEndDate = (date: Date | null | undefined): void => {
        if (date) {
            let endCycle = moment(date)
                .set('hour', moment().hour())
                .set('minute', moment().minute())
                .set('second', moment().second());
            props.getEndDate(endCycle.toDate()!);
        }
        else {
            props.getEndDate(undefined);
        }
    }
    return (
        <>
            {props.screenWidth <= 599 &&
                <div>
                    <Flex gap="gap.small">
                        <Fabric>
                            <Customizer {...datePickerTheme}>
                                <DatePicker
                                    className="date-picker-style-for-small-device"
                                    label={t('startDate')}
                                    placeholder={t('datePlaceholderText')}
                                    isRequired={true}
                                    allowTextInput={true}
                                    showMonthPickerAsOverlay={true}
                                    minDate={props.minStartDate ? new Date(props.minStartDate) : new Date()}
                                    isMonthPickerVisible={true}
                                    value={props.startDate ? new Date(props.startDate) : undefined}
                                    onSelectDate={onSelectStartDate}
                                />
                            </Customizer>
                        </Fabric>
                    </Flex>
                    <Flex gap="gap.small">
                        <Fabric>
                            <Customizer {...datePickerTheme}>
                                <DatePicker
                                    className="date-picker-style-for-small-device"
                                    label={t('endDate')}
                                    placeholder={t('datePlaceholderText')}
                                    isRequired={true}
                                    allowTextInput={true}
                                    minDate={minEndDate}
                                    isMonthPickerVisible={true}
                                    showMonthPickerAsOverlay={true}
                                    value={props.endDate ? new Date(props.endDate) : undefined}
                                    onSelectDate={onSelectEndDate}
                                />
                            </Customizer>
                        </Fabric>
                    </Flex>
                </div>
            }
            {props.screenWidth > 599 &&
                <div>
                    <Flex gap="gap.small">
                        <Flex.Item size="size.half">
                            <div>
                                <Fabric>
                                    <Customizer {...datePickerTheme}>
                                        <DatePicker
                                            className="date-picker-style"
                                            label={t('startDate')}
                                            placeholder={t('datePlaceholderText')}
                                            isRequired={true}
                                            allowTextInput={true}
                                            showMonthPickerAsOverlay={true}
                                            minDate={props.minStartDate ? new Date(props.minStartDate) : new Date()}
                                            isMonthPickerVisible={true}
                                            value={props.startDate ? new Date(props.startDate) : undefined}
                                            onSelectDate={onSelectStartDate}
                                        />
                                    </Customizer>
                                </Fabric>
                            </div>
                        </Flex.Item>
                        <Flex.Item size="size.half">
                            <div>
                                <Fabric>
                                    <Customizer {...datePickerTheme}>
                                        <DatePicker
                                            className="date-picker-style"
                                            label={t('endDate')}
                                            placeholder={t('datePlaceholderText')}
                                            isRequired={true}
                                            allowTextInput={true}
                                            minDate={minEndDate}
                                            isMonthPickerVisible={true}
                                            showMonthPickerAsOverlay={true}
                                            value={props.endDate ? new Date(props.endDate) : undefined}
                                            onSelectDate={onSelectEndDate}
                                        />
                                    </Customizer>
                                </Fabric>
                            </div>
                        </Flex.Item>
                    </Flex>
                </div>
            }
        </>
    );
}
export default StartDateEndDate;