// <copyright file="team-goal-api.ts" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import axios from "./axios-decorator";
import { ITeamGoalDetail, ITeamOwnerDetail } from "../models/type";
import { AxiosResponse } from "axios";

const baseAxiosUrl = window.location.origin;

/**
* Get team goal details by Microsoft Teams' team Id.
* @param teamId {String | Null} Microsoft Teams' team id to fetch specific Team goals.
*/
export const getTeamGoalDetailsByTeamId = async (teamId?: string | null): Promise<AxiosResponse<ITeamGoalDetail[]>> => {
    let url = baseAxiosUrl + `/api/teamgoals?teamId=${teamId}`;
    return await axios.get(url);
}

/**
* Validate if user is team owner.
* @param teamId {String | Null} Microsoft Teams' team id to fetch specific Team goals.
*/
export const getTeamOwnerDetails = async (teamGroupId?: string | null): Promise<AxiosResponse<ITeamOwnerDetail[]>> => {
    let url = baseAxiosUrl + `/api/teamgoals/${teamGroupId}/checkteamowner`;
    return await axios.get(url);
}

/**
* Save team goal details from storage.
* @param teamGoalDetails {Object} Team goal details to be stored in storage.
*/
export const saveTeamGoalDetails = async (teamGoalDetails: {}, teamGroupId?: string | null): Promise<AxiosResponse<void>> => {
    let url = baseAxiosUrl + `/api/teamgoals/${teamGroupId}`;
    return await axios.post(url, teamGoalDetails);
}

/**
* Get specific team goal detail by team goal id.
* @param teamGoalId {String | Null} Team goal id to fetch specific team goal detail.
* @param teamId {String | Null} Microsoft Teams' team id to fetch specific Team goals.
*/
export const getTeamGoalDetailByTeamGoalId = async (teamGoalId?: string | null, teamId?: string | null): Promise<AxiosResponse<ITeamGoalDetail>> => {
    let url = baseAxiosUrl + `/api/teamgoals/goal?teamId=${teamId}&teamGoalId=${teamGoalId}`;
    return await axios.get(url);
}