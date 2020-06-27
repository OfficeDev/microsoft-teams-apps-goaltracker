// <copyright file="personal-goal-api.ts" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import axios from "./axios-decorator";
import { IPersonalGoalDetail } from "../models/type";
import { AxiosResponse } from "axios";

const baseAxiosUrl = window.location.origin;

/**
* Get personal goal details by user AAD object id.
*/
export const getPersonalGoalDetails = async (): Promise<AxiosResponse<IPersonalGoalDetail[]>> => {
    let url = baseAxiosUrl + `/api/personalgoals`;
    return await axios.get(url);
}

/**
* Get personal goal detail by personal goal id.
* @param personalGoalId {String | Null} Unique identifier of personal goal detail entity.
*/
export const getPersonalGoalDetailByGoalIdAsync = async (personalGoalId?: string): Promise<AxiosResponse<IPersonalGoalDetail>> => {

    let url = baseAxiosUrl + `/api/personalgoals/${personalGoalId}`;
    return await axios.get(url);   
}

/**
* Save a personal goal detail in storage.
* @param personalGoalDetail {Object} Personal goal detail to be stored in storage.
*/
export const updatePersonalGoalDetail = async (personalGoalDetail: IPersonalGoalDetail): Promise<AxiosResponse<void>> => {
    let url = baseAxiosUrl + `/api/personalgoals/${personalGoalDetail.PersonalGoalId}`;
    return await axios.patch(url, personalGoalDetail);
}

/**
* Save all personal goal details in storage.
* @param personalGoalDetails {Object} Personal goal details to be stored in storage.
*/
export const savePersonalGoalDetails = async (personalGoalDetails: {}): Promise<AxiosResponse<void>> => {
    let url = baseAxiosUrl + "/api/personalgoals";
    return await axios.post(url, personalGoalDetails);    
}

/**
* delete specified personal goal detail.
* @param personalGoalDetails {Object} Personal goal detail to be deleted from storage.
*/
export const deletePersonalGoalDetail = async (personalGoalId: {}): Promise<AxiosResponse<void>> => {
    let url = baseAxiosUrl + `/api/personalgoals/${personalGoalId}`;
    return await axios.delete(url);
}