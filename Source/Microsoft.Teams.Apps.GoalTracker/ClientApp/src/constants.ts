/*
    <copyright file="constants.ts" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/

export default class Constants {

	//Themes
	public static readonly body: string = "body";
	public static readonly theme: string = "theme";
	public static readonly default: string = "default";
	public static readonly light: string = "light";
	public static readonly dark: string = "dark";
    public static readonly contrast: string = "contrast";

	public static readonly maxAllowedGoalName = 300;
	public static readonly maxAllowedNoteDescription = 1000;
	public static readonly maxAllowedGoals = 15;
	public static readonly setPersonalGoal: string ="set personal goals";
	public static readonly editPersonalGoal: string = "edit personal goals";
	public static readonly editTeamGoal: string = "edit team goals";
	public static readonly setTeamGoal: string = "set team goals";

	// Date formats
	public static readonly goalCycleDateTimeFormat = "ll"; // This format will be used to display goal cycles dates as per user's locale on UI.
	public static readonly dateComparisonFormat = "YYYY-MM-DD";
	public static readonly dateTimeOffsetFormat = "YYYY-MM-DD HH:mm Z";
	public static readonly utcDateFormat = "MM-DD-YYYY";
}

