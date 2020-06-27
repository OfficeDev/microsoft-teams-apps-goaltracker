// <copyright file="goal-helper.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

export const getGoalStatusCollection = (t: any) => {
	return [
		{
			DisplayName: t('notStartedStatus'),
			value: 0,
			header: t('notStartedStatus')
		},
		{
			DisplayName: t('inProgressStatus'),
			value: 1,
			header: t('inProgressStatus')
		},
		{
			DisplayName: t('completedStatus'),
			value: 2,
			header: t('completedStatus')
		}
	]
}

export const getReminderFrequncyCollection = (t: any, isReminderActive: boolean) => {
	return [
		{
			name: t('weeklyReminderFrequency'),
			key: t('weeklyReminderFrequency'),
			label: t('weeklyReminderFrequency'),
			value: 0,
			disabled: isReminderActive,
		},
		{
			name: t('biweeklyReminderFrequency'),
			key: t('biweeklyReminderFrequency'),
			label: t('biweeklyReminderFrequency'),
			value: 1,
			disabled: isReminderActive,
		},
		{
			name: t('monthlyReminderFrequency'),
			key: t('monthlyReminderFrequency'),
			label: t('monthlyReminderFrequency'),
			value: 2,
			disabled: isReminderActive,
		},
		{
			name: t('quarterlyReminderFrequency'),
			key: t('quarterlyReminderFrequency'),
			label: t('quarterlyReminderFrequency'),
			value: 3,
			disabled: isReminderActive,
		},
	];
}

/**
* Handle error occurred during API call.
* @param error {Object} Error response object.
*/
export const handleError = (error: any): any => {
	const errorStatus = error.status;
	if (errorStatus === 403) {
		window.location.href = `/error?code=403`;
	}
	else if (errorStatus === 401) {
		window.location.href = `/error?code=401`;
	}
	else {
		window.location.href = `/error`;
	}
}