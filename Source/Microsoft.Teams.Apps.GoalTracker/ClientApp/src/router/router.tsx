// <copyright file="router.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>


import * as React from "react";
import { BrowserRouter, Route, Switch } from "react-router-dom";
import { Suspense } from "react";
import { Loader } from "@fluentui/react-northstar";
import ErrorPage from '../components/error-page';
import ManageGoals from '../components/manage-goals/manage-goals';
import EditGoal from "../components/manage-goals/edit-goal-detail";
import SignInPage from "../components/signin/signin";
import SignInSimpleStart from "../components/signin/signin-start";
import SignInSimpleEnd from "../components/signin/signin-end";
import PersonalGoal from '../components/setgoals/personal-goal'
import TeamGoal from '../components/setgoals/team-goal'
import AlignGoal from '../components/align-goal/align-goal';

export const AppRoute: React.FunctionComponent<{}> = () => {
	return (
		<Suspense fallback={<div> <Loader /></div>}>
			<BrowserRouter>
				<Switch>
					<Route exact path="/" component={ManageGoals} />
					<Route exact path="/error" component={ErrorPage} />
					<Route exact path="/manage-goals" component={ManageGoals} />
					<Route exact path="/signin" component={SignInPage} />
					<Route exact path="/signin-simple-start" component={SignInSimpleStart} />
                    <Route exact path="/signin-simple-end" component={SignInSimpleEnd} />
                    <Route exact path="/personal-goal" component={PersonalGoal} />
					<Route exact path="/edit-goal-detail" component={EditGoal} />
                    <Route exact path="/team-goal" component={TeamGoal} />
					<Route exact path="/align-goal" component={AlignGoal} />
				</Switch>
			</BrowserRouter>
		</Suspense>
	);
};