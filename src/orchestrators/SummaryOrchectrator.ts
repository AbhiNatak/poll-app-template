import { Logger } from './../utils/Logger';
import { Constants } from '../utils/Constants';
import { Localizer } from '../utils/Localizer';
import { initialize, setProgressStatus, setContext, fetchLocalization, fetchUserDetails, fetchActionInstance, fetchActionInstanceSummary, fetchMyResponse, fetchMemberCount, setIsActionDeleted, updateMyRow, updateActionInstance, fetchActionInstanceRows, updateMemberCount, updateActionInstanceSummary, addActionInstanceRows, updateContinuationToken, updateNonResponders, closePoll, fetchNonReponders, pollCloseAlertOpen, deletePoll, pollDeleteAlertOpen, updateDueDate, pollExpiryChangeAlertOpen, downloadCSV, updateUserProfileInfo } from '../actions/SummaryActions';
import { orchestrator } from 'satcheljs';
import { ProgressState } from '../utils/SharedEnum';
import getStore from "../store/SummaryStore";
import * as actionSDK from "@microsoft/m365-action-sdk";
import { ActionSdkHelper } from "../helper/ActionSdkHelper";

const handleErrorResponse = (error: actionSDK.ApiError) => {
    if (error && error.code == "404") {
        setIsActionDeleted(true);
    }
};

const handleError = (error: actionSDK.ApiError, requestType: string) => {
    handleErrorResponse(error);
    setProgressStatus({ [requestType]: ProgressState.Failed });
}

orchestrator(initialize, async () => {
    if (
        getStore().progressStatus.currentContext == ProgressState.NotStarted ||
        getStore().progressStatus.currentContext == ProgressState.Failed
    ) {
        setProgressStatus({ currentContext: ProgressState.InProgress });

        let actionContext = await ActionSdkHelper.getActionContext();
        if (actionContext.success) {
            let context = actionContext.context as actionSDK.ActionSdkContext
            setContext(context);
            fetchLocalization();
            fetchUserDetails([context.userId]);
            fetchActionInstance(true);
            fetchActionInstanceSummary(true);
            fetchMyResponse();
            fetchMemberCount();
            setProgressStatus({ currentContext: ProgressState.Completed });
        } else {
            handleError(actionContext.error, "currentContext");
        }
    }
});

orchestrator(fetchLocalization, async (msg) => {
    if (
        getStore().progressStatus.localizationState == ProgressState.NotStarted ||
        getStore().progressStatus.localizationState == ProgressState.Failed
    ) {
        setProgressStatus({ localizationState: ProgressState.InProgress });
        let response = await Localizer.initialize();
        response ? setProgressStatus({ localizationState: ProgressState.Completed }) : setProgressStatus({ localizationState: ProgressState.Failed });
    }
});

orchestrator(fetchMyResponse, async () => {
    if (
        getStore().progressStatus.myActionInstanceRow == ProgressState.NotStarted ||
        getStore().progressStatus.myActionInstanceRow == ProgressState.Failed
    ) {
        setProgressStatus({ myActionInstanceRow: ProgressState.InProgress });

        let response = await ActionSdkHelper.getActionDataRows(getStore().context.actionId, "self", null, 1);

        if (response.success) {
            let row: actionSDK.ActionDataRow = response.dataRows[0];
            updateMyRow(row);
            setProgressStatus({ myActionInstanceRow: ProgressState.Completed });
        } else {
            handleError(response.error, "myActionInstanceRow");
        }
    }
});

orchestrator(fetchMemberCount, async (msg) => {
    if (
        getStore().progressStatus.memberCount == ProgressState.NotStarted ||
        getStore().progressStatus.memberCount == ProgressState.Failed
    ) {
        setProgressStatus({ memberCount: ProgressState.InProgress });

        let response = await ActionSdkHelper.getSubscriptionMemberCount(getStore().context.subscription);
        if (response.success) {
            updateMemberCount(response.memberCount);
            setProgressStatus({ memberCount: ProgressState.Completed });
        } else {
            handleError(response.error, "memberCount");
        }
    }
});

orchestrator(fetchActionInstance, async (msg) => {
    if (getStore().progressStatus.actionInstance != ProgressState.InProgress) {
        if (msg.updateProgressState) {
            setProgressStatus({ actionInstance: ProgressState.InProgress });
        }
        let response = await ActionSdkHelper.getAction(getStore().context.actionId);
        if (response.success) {
            updateActionInstance(response.action);
            fetchActionInstanceRows(false);
            msg.updateProgressState && setProgressStatus({ actionInstance: ProgressState.Completed });
        } else {
            msg.updateProgressState && setProgressStatus({ actionInstance: ProgressState.Failed });
        }
    }
});

orchestrator(fetchActionInstanceSummary, async (msg) => {
    if (
        getStore().progressStatus.actionInstanceSummary != ProgressState.InProgress
    ) {
        if (msg.updateProgressState) {
            setProgressStatus({ actionInstanceSummary: ProgressState.InProgress });
        }

        let response = await ActionSdkHelper.getActionDataRowsSummary(getStore().context.actionId, true);
        if (response.success) {
            updateActionInstanceSummary(response.summary);
            msg.updateProgressState && setProgressStatus({ actionInstanceSummary: ProgressState.Completed });
        } else {
            msg.updateProgressState && setProgressStatus({ actionInstanceSummary: ProgressState.Failed });
        }
    }
});

orchestrator(fetchUserDetails, async (msg) => {
    let userIds: string[] = [];

    // fetch only those user profiles that are not present in the store
    for (var userId of msg.userIds) {
        if (
            !getStore().userProfile.hasOwnProperty(userId) ||
            !getStore().userProfile[userId].displayName
        ) {
            userIds.push(userId);
        }
    }
    if (userIds.length > 0) {

        let response = await ActionSdkHelper.getSubscriptionMembers(getStore().context.subscription, userIds);

        if (response.success && response.members) {
            let users: {
                [key: string]: actionSDK.SubscriptionMember;
            } = {};
            response.members.forEach(member => {
                users[member.id] = { id: member.id, displayName: member.displayName }
            });
            updateUserProfileInfo(users);
            if (response.memberIdsNotFound) {
                let userProfile: {
                    [key: string]: actionSDK.SubscriptionMember;
                } = {};
                for (var userId of response.memberIdsNotFound) {
                    userProfile[userId] = { id: userId, displayName: null };
                }
                updateUserProfileInfo(userProfile);
            }
        } else if (!response.success) {
            handleErrorResponse(response.error);
            let userProfile: { [key: string]: actionSDK.SubscriptionMember } = {};
            for (var userId of userIds) {
                userProfile[userId] = { id: userId, displayName: null };
            }
            updateUserProfileInfo(userProfile);
        }
    }
});

orchestrator(fetchActionInstanceRows, async (msg) => {
    if (
        getStore().actionInstance &&
        (getStore().actionInstance.dataTables[0].rowsVisibility ==
            actionSDK.Visibility.All ||
            (getStore().actionInstance.dataTables[0].rowsVisibility ==
                actionSDK.Visibility.Sender &&
                getStore().actionInstance.creatorId == getStore().context.userId))
    ) {
        if (
            getStore().progressStatus.actionInstanceRow == ProgressState.Partial ||
            getStore().progressStatus.actionInstanceRow == ProgressState.Failed ||
            getStore().progressStatus.actionInstanceRow == ProgressState.NotStarted
        ) {
            setProgressStatus({ actionInstanceRow: ProgressState.InProgress });

            let response = await ActionSdkHelper.getActionDataRows(getStore().context.actionId, null, getStore().continuationToken, 30);

            if (response.success) {
                let rows: actionSDK.ActionDataRow[] = [];
                for (var row of response.dataRows) {
                    rows.push(row);
                }

                let userIds: string[] = [];
                for (var row of rows) {
                    userIds.push(row.creatorId);
                }

                addActionInstanceRows(rows);
                if (msg.shouldFetchUserDetails) {
                    fetchUserDetails(userIds);
                }
                if (response.continuationToken) {
                    updateContinuationToken(response.continuationToken);
                    setProgressStatus({ actionInstanceRow: ProgressState.Partial });
                } else {
                    setProgressStatus({ actionInstanceRow: ProgressState.Completed });
                }
            } else {
                handleError(response.error, "actionInstanceRow");
            }
        }
    }
});

orchestrator(fetchNonReponders, async () => {
    if (
        getStore().progressStatus.nonResponder == ProgressState.NotStarted ||
        getStore().progressStatus.nonResponder == ProgressState.Failed
    ) {
        setProgressStatus({ nonResponder: ProgressState.InProgress });


        let response = await ActionSdkHelper.getNonResponders(getStore().context.actionId, getStore().context.subscription.id);

        if (response.success) {
            let userProfile: { [key: string]: actionSDK.SubscriptionMember } = {};
            response.nonParticipants.forEach(
                (user: actionSDK.SubscriptionMember) => {
                    userProfile[user.id] = user;
                }
            );
            updateUserProfileInfo(userProfile);
            updateNonResponders(response.nonParticipants);
            setProgressStatus({ nonResponder: ProgressState.Completed });
        } else {
            handleError(response.error, "nonResponder");
        }
    }
});

orchestrator(closePoll, async () => {
    if (
        getStore().progressStatus.closeActionInstance != ProgressState.InProgress
    ) {
        let failedCallback = () => {
            setProgressStatus({ closeActionInstance: ProgressState.Failed });
            fetchActionInstance(false);
        };

        setProgressStatus({ closeActionInstance: ProgressState.InProgress });
        var actionInstanceUpdateInfo: actionSDK.ActionUpdateInfo = {
            id: getStore().context.actionId,
            version: getStore().actionInstance.version,
            status: actionSDK.ActionStatus.Closed,
        };

        let response = await ActionSdkHelper.updateActionInstance(actionInstanceUpdateInfo);
        if (response.success) {
            if (response.updateSuccess) {
                pollCloseAlertOpen(false);
                await ActionSdkHelper.closeView();
            } else {
                Logger.logError(`closePoll failed, Error: not success`);
                failedCallback();
            }
        } else {
            handleErrorResponse(response.error);
            failedCallback();
        }
    }
});

orchestrator(deletePoll, async () => {
    if (
        getStore().progressStatus.deleteActionInstance != ProgressState.InProgress
    ) {
        let failedCallback = () => {
            setProgressStatus({ deleteActionInstance: ProgressState.Failed });
            fetchActionInstance(false);
        };

        setProgressStatus({ deleteActionInstance: ProgressState.InProgress });

        let response = await ActionSdkHelper.deleteActionInstance(getStore().context.actionId);
        if (response.success) {
            if (response.deleteSuccess) {
                pollDeleteAlertOpen(false);
                await ActionSdkHelper.closeView();
            } else {
                Logger.logError(`deletePoll failed, Error: not success`);
                failedCallback();
            }
        } else {
            handleErrorResponse(response.error);
            failedCallback();
        }
    }
});

orchestrator(updateDueDate, async (actionMessage) => {
    if (
        getStore().progressStatus.updateActionInstance != ProgressState.InProgress
    ) {
        let callback = (success: boolean) => {
            setProgressStatus({
                updateActionInstance: success
                    ? ProgressState.Completed
                    : ProgressState.Failed,
            });
            fetchActionInstance(false);
        };

        setProgressStatus({ updateActionInstance: ProgressState.InProgress });
        var actionInstanceUpdateInfo: actionSDK.ActionUpdateInfo = {
            id: getStore().context.actionId,
            version: getStore().actionInstance.version,
            expiryTime: actionMessage.dueDate,
        };

        let response = await ActionSdkHelper.updateActionInstance(actionInstanceUpdateInfo);
        if (response.success) {
            if (response.updateSuccess) {
                callback(true);
                pollExpiryChangeAlertOpen(false);
            } else {
                Logger.logError(`updateDueDate failed, Error: not success`);
                callback(false);
            }
        } else {
            handleErrorResponse(response.error);
            callback(false);
        }
    }
});

orchestrator(downloadCSV, async (msg) => {
    if (getStore().progressStatus.downloadData != ProgressState.InProgress) {
        setProgressStatus({ downloadData: ProgressState.InProgress });

        let response = await ActionSdkHelper.downloadCSV(getStore().context.actionId,
            Localizer.getString(
                "PollResult",
                getStore().actionInstance.dataTables[0].dataColumns[0].displayName
            ).substring(0, Constants.ACTION_RESULT_FILE_NAME_MAX_LENGTH)
        );

        response.success ?
            setProgressStatus({ downloadData: ProgressState.Completed }) : handleError(response.error, "downloadData")
    }
});