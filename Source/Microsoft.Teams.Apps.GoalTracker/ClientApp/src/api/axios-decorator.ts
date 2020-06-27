// <copyright file="axios-decorator.ts" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import axios, { AxiosResponse, AxiosRequestConfig } from "axios";
import * as microsoftTeams from "@microsoft/teams-js";

export class AxiosJWTDecorator {

    /**
	* Delete data
	* @param  {String} url Resource URI
	*/
    public async delete<T = any, R = AxiosResponse<T>>(
        url: string,
        data?: any,
        config?: AxiosRequestConfig,
        needAuthorizationHeader: boolean = true
    ): Promise<R> {
        try {
            let config: AxiosRequestConfig = axios.defaults;
            if (needAuthorizationHeader) {
                config = await this.setupAuthorizationHeader(config);
            }
            if (data) {
                config.headers["Content-Type"] = 'application/json; charset=utf-8';
                config.data = data;
            }

            return await axios.delete(url, config);
        } catch (error) {
            return error.response;
        }
    }

	/**
	* Post data to API
	* @param  {String} url Resource URI
	* @param  {Object} data Request body data
	*/
    public async post<T = any, R = AxiosResponse<T>>(
        url: string,
        data?: any,
        config?: AxiosRequestConfig,
        needAuthorizationHeader: boolean = true
    ): Promise<R> {
        try {
            let config: AxiosRequestConfig = axios.defaults;
            if (needAuthorizationHeader) {
                config = await this.setupAuthorizationHeader(config);
            }

            return await axios.post(url, data, config);
        } catch (error) {
            return error.response;
        }
    }

	/**
	* Update data
	* @param  {String} url Resource URI
	* @param  {Object} data Request body data
	*/
    public async put<T = any, R = AxiosResponse<T>>(
        url: string,
        data?: any,
        config?: AxiosRequestConfig,
        needAuthorizationHeader: boolean = true
    ): Promise<R> {
        try {
            if (needAuthorizationHeader) {
                config = await this.setupAuthorizationHeader(config);
            }

            return await axios.put(url, data, config);
        } catch (error) {
            return error.response;
        }
    }

    /**
	* Update data with patch request
	* @param  {String} url Resource URI
	* @param  {Object} data Request body data
	*/
    public async patch<T = any, R = AxiosResponse<T>>(
        url: string,
        data?: any,
        config?: AxiosRequestConfig,
        needAuthorizationHeader: boolean = true
    ): Promise<R> {
        try {
            if (needAuthorizationHeader) {
                config = await this.setupAuthorizationHeader(config);
            }

            return await axios.patch(url, data, config);
        } catch (error) {
            return error.response;
        }
    }

	/**
	* Get data from API
	*/
    public async get<T = any, R = AxiosResponse<T>>(
        url: string,
        config?: AxiosRequestConfig,
        needAuthorizationHeader: boolean = true
    ): Promise<R> {
        try {
            if (needAuthorizationHeader) {
                config = await this.setupAuthorizationHeader(config);
            }

            return await axios.get(url, config);
        } catch (error) {
            return error.response;
        }
    }

    /**
    * Sets authorization header in request.
    */
    private async setupAuthorizationHeader(
        config?: AxiosRequestConfig
    ): Promise<AxiosRequestConfig> {
        microsoftTeams.initialize();

        return new Promise<AxiosRequestConfig>((resolve, reject) => {
            const authTokenRequest = {
                successCallback: (token: string) => {
                    if (!config) {
                        config = axios.defaults;
                    }
                    config.headers["Authorization"] = `Bearer ${token}`;
                    resolve(config);
                },
                failureCallback: (error: string) => {
                    // When the getAuthToken function returns a "resourceRequiresConsent" error, 
                    // it means Azure AD needs the user's consent before issuing a token to the app. 
                    // The following code redirects the user to the "Sign in" page where the user can grant the consent. 
                    // Right now, the app redirects to the consent page for any error.
                    console.error("Error from getAuthToken: ", error);
                    window.location.href = "/signin";
                },
                resources: []
            };
            microsoftTeams.authentication.getAuthToken(authTokenRequest);
        });
    }
}

const axiosJWTDecoratorInstance = new AxiosJWTDecorator();
export default axiosJWTDecoratorInstance;