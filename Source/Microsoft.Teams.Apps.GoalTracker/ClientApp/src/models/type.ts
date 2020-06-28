// <copyright file="type.ts" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

export interface ITeamGoalDetail {
    CreatedOn: string,
    CreatedBy?: string | null,
    LastModifiedOn: string,
    LastModifiedBy?: string | null,
    IsActive: boolean,
    IsDeleted: boolean,
    ReminderFrequency: number,
    IsReminderActive: boolean,
    TeamId?: string | null,
    TeamGoalId: string,
    TeamGoalName: string,
    TeamGoalStartDate: string,
    TeamGoalEndDate: string,
    TeamGoalEndDateUTC: string,
    AdaptiveCardActivityId: string,
    ServiceURL: string | null,
    GoalCycleId: string,
}
export interface IPersonalGoalDetail {
    UserAadObjectId?: string | null,
    AdaptiveCardActivityId: string,
    ConversationId?: string | null,
    IsActive: boolean,
    IsAligned: boolean,
    IsDeleted: boolean,
    CreatedOn: string,
    CreatedBy?: string | null,
    LastModifiedOn: string,
    LastModifiedBy?: string | null,
    GoalName: string,
    PersonalGoalId: string,
    ReminderFrequency: number,
    IsReminderActive: boolean,
    Status: number,
    StartDate: string,
    EndDate: string,
    ServiceURL: string | null,
    TeamId?: string | null,
    TeamGoalId?: string | null,
    TeamGoalName?: string | null,
    EndDateUTC: string,
    NotesCount: number,
    GoalCycleId: string,
}
export interface IPersonalGoalNoteDetail {
    CreatedOn: string,
    CreatedBy: string,
    LastModifiedOn: string,
    LastModifiedBy: string,
    IsActive: boolean,
    PersonalGoalId: string,
    PersonalGoalNoteId: string,
    PersonalGoalNoteDescription: string,
    SourceName: string,
    UserAadObjectId: string,
    NotesCount: number,
    IsEdited: boolean
}
export interface IAddNewGoal {
    key: string
    header: JSX.Element,
    goalName: string
}

export interface ITeamOwnerDetail {
    TeamOwnerId: string
}