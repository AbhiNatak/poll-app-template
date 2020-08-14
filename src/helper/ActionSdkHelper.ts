import * as actionSDK from "@microsoft/m365-action-sdk";
import { Logger } from './../utils/Logger';

export class ActionSdkHelper {

    /**
     * API to fetch current action context 
     */
    public static async getActionContext() {
        try {
            let response = await actionSDK.executeApi(new actionSDK.GetContext.Request()) as actionSDK.GetContext.Response;
            Logger.logInfo("fetchCurrentContext - Response: " + JSON.stringify(response));
            return { success: true, context: response.context };
        } catch (error) {
            Logger.logError(`fetchCurrentContext failed, Error: ${error.category}, ${error.code}, ${error.message}`);
            return { success: false, error: error };
        }
    }

    /*
    * @desc Service Request to create new Action Instance 
    * @param {actionSDK.Action} action instance which need to get created
    */
    public static async createActionInstance(action: actionSDK.Action) {
        let request = new actionSDK.CreateAction.Request(action);
        try {
            let response = await actionSDK.executeApi(request) as actionSDK.GetContext.Response;
            Logger.logInfo("createActionInstance - Response: " + JSON.stringify(response));
        } catch (error) {
            Logger.logError(`createActionInstance failed, Error: ${error.category}, ${error.code}, ${error.message}`);
        }
    }

    /**
     * Function to get for data rows for a user 
     * @param actionId action instance id
     * @param createrId created id
     * @param continuationToken 
     * @param pageSize 
     */
    public static async getActionDataRows(actionId: string, creatorId?: string, continuationToken?: string, pageSize?: number) {
        let request = new actionSDK.GetActionDataRows.Request(actionId, creatorId, continuationToken, pageSize);
        try {
            let response = await actionSDK.executeApi(request) as actionSDK.GetActionDataRows.Response;
            return { success: true, dataRows: response.dataRows, continuationToken: response.continuationToken }
        } catch (error) {
            Logger.logError(`fetchMyResponse failed, Error: ${error.category}, ${error.code}, ${error.message}`);
            return { success: false, error: error }
        }
    }

    /*
    *   @desc Service API Request for getting the membercount
    *   @param subscription - action subscription: actionSDK.ActionSdkContext.subscription
    */
    public static async getSubscriptionMemberCount(subscription: actionSDK.Subscription) {
        let getSubscriptionCount = new actionSDK.GetSubscriptionMemberCount.Request(subscription);
        try {
            let response = await actionSDK.executeApi(getSubscriptionCount) as actionSDK.GetSubscriptionMemberCount.Response;
            return { success: true, memberCount: response.memberCount };
        } catch (error) {
            Logger.logError(`fetchMemberCount failed, Error: ${error.category}, ${error.code}, ${error.message}`);
            return { success: false, error: error };
        }
    }

    /*
    * @desc Service API Request for fetching action instance
    * @param {actionId} action id for which we want to get details
    */
    public static async getAction(actionId?: string) {
        try {
            let response = await actionSDK.executeApi(new actionSDK.GetAction.Request(actionId)) as actionSDK.GetAction.Response;
            Logger.logInfo("fetchActionInstance success, response: " + JSON.stringify(response));
            return { success: true, action: response.action }
        } catch (error) {
            Logger.logError(`fetchActionInstance failed, Error: ${error.category}, ${error.code}, ${error.message}`);
            return { success: false, error: error }
        }
    }

    /**
     * Funtion to get action data summary 
     * @param actionId action id 
     * @param addDefaultAggregates 
     */
    public static async getActionDataRowsSummary(actionId: string, addDefaultAggregates?: boolean) {
        let request = new actionSDK.GetActionDataRowsSummary.Request(actionId, addDefaultAggregates);
        try {
            let response = await actionSDK.executeApi(request) as actionSDK.GetActionDataRowsSummary.Response;
            return { success: true, summary: response.summary }
        } catch (error) {
            Logger.logError(`fetchActionInstanceSummary failed, Error: ${error.category}, ${error.code}, ${error.message}`);
            return { success: false, error: error }
        }
    }

    /**
     * Method to get details of member of subscription
     * @param subscription subscription
     * @param userId user id to get details
     */
    public static async getSubscriptionMembers(subscription, userIds) {
        try {
            let request = new actionSDK.GetSubscriptionMembers.Request(subscription, userIds);
            let response = await actionSDK.executeApi(request) as actionSDK.GetSubscriptionMembers.Response;
            return { success: true, members: response.members, memberIdsNotFound: response.memberIdsNotFound }
        } catch (error) {
            Logger.logError(`fetchUserDetails failed, Error: ${error.category}, ${error.code}, ${error.message}`);
            return { success: false, error: error }
        }
    }

    /**
     * @desc Service API Request for getting the nonResponders details
     * @param actionId actionId
     * @param subscriptionId subscriptionId
     */
    public static async getNonResponders(actionId: string, subscriptionId: string) {
        let requestNonResponders = new actionSDK.GetActionSubscriptionNonParticipants.Request(actionId, subscriptionId);
        try {
            let response = await actionSDK.executeApi(requestNonResponders) as actionSDK.GetActionSubscriptionNonParticipants.Response;
            return { success: true, nonParticipants: response.nonParticipants }
        } catch (error) {
            Logger.logError(`fetchNonReponders failed, Error: ${error.category}, ${error.code}, ${error.message}`);
            return { sucess: false, error: error }
        }
    }

    /**
     * Method to update action instance data 
     * @param data object of data we want modify
    */
    public static async updateActionInstance(actionUpdateInfo: actionSDK.ActionUpdateInfo) {
        let getUpdateActionRequest = new actionSDK.UpdateAction.Request(actionUpdateInfo);
        try {
            let response = await actionSDK.executeApi(getUpdateActionRequest) as actionSDK.UpdateAction.Response;
            Logger.logInfo("UpdateAction Success, Response: " + JSON.stringify(response));
            return { success: true, updateSuccess: response.success }
        } catch (error) {
            Logger.logError(`UpdateAction failed, Error: ${error.category}, ${error.code}, ${error.message}`);
            return { success: false, error: error }
        }
    }

    /**
     * API to close current view
     */
    public static async closeView() {
        let closeViewRequest = new actionSDK.CloseView.Request();
        await actionSDK.executeApi(closeViewRequest);
    }

    /**
     * Method to delete action instance 
     * @param actionId action instance id
     */
    public static async deleteActionInstance(actionId) {
        try {
            let request = new actionSDK.DeleteAction.Request(actionId);
            let response = await actionSDK.executeApi(request) as actionSDK.DeleteAction.Response;
            return { success: true, deleteSuccess: response.success }
        } catch (error) {
            Logger.logError(`deletePoll failed, Error: ${error.category}, ${error.code}, ${error.message}`);
            return { success: false, error: error }
        }
    }

    public static async downloadCSV(actionId, fileName) {
        let request = new actionSDK.DownloadActionDataRowsResult.Request(actionId, fileName);
        try {
            let response = actionSDK.executeApi(request);
            return { success: true }
        } catch (error) {
            Logger.logError(`downloadCSV failed, Error: ${error.category}, ${error.code}, ${error.message}`);
            return { success: false, error: error }
        }
    }

    /*
    * @desc Gets the localized strings in which the app is rendered
    */
    public static async getLocalizedStrings() {
        let request = new actionSDK.GetLocalizedStrings.Request()
        try {
            let response = await actionSDK.executeApi(request) as actionSDK.GetLocalizedStrings.Response;
            return { success: true, strings: response.strings }
        } catch (error) {
            Logger.logError(`fetchLocalization failed, Error: ${error.category}, ${error.code}, ${error.message}`);
        }
    }

    /**
     * Method to hide loading indicater
     */
    public static hideLoadingIndicator() {
        actionSDK.executeApi(new actionSDK.HideLoadingIndicator.Request());
    }
}
