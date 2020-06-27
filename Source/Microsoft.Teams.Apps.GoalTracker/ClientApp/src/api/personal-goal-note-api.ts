/*
    <copyright file="personal-goal-note-api.ts" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/

import axios from "./axios-decorator";
import { IPersonalGoalNoteDetail } from "../models/type";
import { AxiosResponse } from "axios";

const baseAxiosUrl = window.location.origin;

/**
* Get personal goal note details by user Azure Active Directory object id.
*/
export const getPersonalGoalNotesCount = async (): Promise<AxiosResponse<IPersonalGoalNoteDetail[]>> => {

    let url = baseAxiosUrl + `/api/notes/count`;
    return await axios.get(url);
}

/**
* Get personal goal details by personal goal id.
* @param personalGoalId {String | Null} Unique identifier of personal goal detail entity.
*/
export const getPersonalGoalNoteDetails = async (personalGoalId?: string): Promise<AxiosResponse<IPersonalGoalNoteDetail[]>> => {

    let url = baseAxiosUrl + `/api/notes/goal/${personalGoalId}`;
    return await axios.get(url);
}

/**
* Save personal goal note details from storage.
* @param personalGoalNoteDetails {Object} Personal goal note details to be stored in storage.
*/
export const savePersonalGoalNoteDetails = async (personalGoalNoteDetails: {}): Promise<AxiosResponse<void>> => {
    let url = baseAxiosUrl + "/api/notes";
    return await axios.put(url, personalGoalNoteDetails);
}

/**
* Delete personal goal note details from storage.
* @param personalGoalNoteIds {Object} Collection of personal goal note ids to be deleted from storage.
*/
export const deletePersonalGoalNoteDetails = async (personalGoalNoteIds: {}): Promise<AxiosResponse<void>> => {
    let url = baseAxiosUrl + "/api/notes";
    return await axios.delete(url, personalGoalNoteIds);
}