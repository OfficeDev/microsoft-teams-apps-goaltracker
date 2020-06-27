﻿// <copyright file="people-picker-styles.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import { IPeoplePickerItemSelectedStyleProps, IPeoplePickerItemSelectedStyles } from 'office-ui-fabric-react';
import { DarkTheme } from './dark-customizations';

export const PeoplePickerItemStyles = (
    props: IPeoplePickerItemSelectedStyleProps,
): Partial<IPeoplePickerItemSelectedStyles> => {
    const { selected } = props;

    return {
        root: [
            {
                background: DarkTheme.palette.neutralQuaternaryAlt,
                selectors: {
                    ':hover': {
                        background: DarkTheme.palette.neutralQuaternary,
                    },
                },
            },
            selected && {
                backgroundColor: DarkTheme.palette.themeSecondary,
                selectors: {
                    ':hover': {
                        background: DarkTheme.palette.themeSecondary,
                    },
                },
            },
        ],
        removeButton: [
            {
                background: DarkTheme.palette.neutralQuaternaryAlt,
                color: DarkTheme.palette.neutralDark,
                selectors: {
                    ':hover': {
                        background: DarkTheme.palette.neutralQuaternary,
                        color: DarkTheme.palette.black,
                    },
                },
            },
            selected && {
                background: DarkTheme.palette.themeSecondary,
                selectors: {
                    ':hover': {
                        background: DarkTheme.palette.themeDark,
                    },
                    ':active': {
                        color: DarkTheme.palette.black,
                    },
                },
            },
        ],
        subComponentStyles: {
            persona: {
                primaryText: [
                    {
                        color: DarkTheme.palette.neutralPrimary,
                    },
                    selected && {
                        color: DarkTheme.palette.black,
                        selectors: {
                            ':hover': {
                                color: DarkTheme.palette.black,
                            },
                        },
                    },
                ],
            },
        },
    };
};
