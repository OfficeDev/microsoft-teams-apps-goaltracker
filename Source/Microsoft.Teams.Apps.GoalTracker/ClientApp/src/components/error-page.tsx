// <copyright file="error-page.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import * as React from "react";
import { Text, Flex, Label } from "@fluentui/react-northstar";
import { ErrorIcon } from '@fluentui/react-icons-northstar';
import { WithTranslation, withTranslation } from "react-i18next";
import { TFunction } from "i18next";

class ErrorPage extends React.Component<WithTranslation, {}> {
    code: string | null = null;
    localize: TFunction;

    constructor(props: any) {
        super(props);
        let search = window.location.search;
        let params = new URLSearchParams(search);
        this.code = params.get("code");
        this.localize = this.props.t;
        this.state = {
            resourceStrings: {}
        };
    }

    /** Called once component is mounted. */
    async componentDidMount() {
    }

    /**
     * Render error page.
     * */
    render() {
        let message = this.localize("genericErrorMessage");
        if (this.code === "401") {
            message = `${this.localize("unauthorizedAccessMessage")}`;
        }

        return (
            <div className="container-div">
                <Flex gap="gap.small" hAlign="center" vAlign="center" className="error-container">
                    <Flex gap="gap.small" hAlign="center" vAlign="center">
                        <Flex.Item>
                            <div
                                style={{
                                    position: "relative",
                                }}
                            >
                                <Label icon={<ErrorIcon />} />
                            </div>
                        </Flex.Item>

                        <Flex.Item grow>
                            <Flex column gap="gap.small" vAlign="stretch">
                                <div>
                                    <Text weight="bold" error content={message} /><br />
                                </div>
                            </Flex>
                        </Flex.Item>
                    </Flex>
                </Flex>
            </div>
        );
    }
}

export default withTranslation()(ErrorPage);