import { Logger } from './../utils/Logger';
import { toJS } from 'mobx';
import { Localizer } from '../utils/Localizer';
import { orchestrator } from 'satcheljs';
import { setContext, initialize, setAppInitialized, callActionInstanceCreationAPI, updateTitle, updateChoiceText, setSendingFlag, shouldValidateUI } from '../actions/CreationActions';
import { fetchCurrentContext } from "../actions/CreationActions";
import { ProgressState } from '../utils/SharedEnum';
import getStore from "../store/CreationStore";
import { Utils } from '../utils/Utils';
import * as actionSDK from "@microsoft/m365-action-sdk";
import { ActionSdkHelper } from "../helper/ActionSdkHelper";

function validateActionInstance(actionInstance: actionSDK.Action): boolean {
    if (actionInstance == null) return false;
    if (
        actionInstance.dataTables[0].dataColumns == null ||
        actionInstance.dataTables[0].dataColumns.length <= 0
    )
        return false;
    if (
        actionInstance.dataTables[0].dataColumns[0].displayName == null ||
        actionInstance.dataTables[0].dataColumns[0].displayName == ""
    )
        return false;
    if (actionInstance.dataTables[0].dataColumns[0].options == null) return false;
    if (actionInstance.dataTables[0].dataColumns[0].options.length < 2)
        return false;
    for (var option of actionInstance.dataTables[0].dataColumns[0].options) {
        if (option.displayName == null || option.displayName == "") {
            return false;
        }
    }
    return true;
}

orchestrator(fetchCurrentContext, async () => {
    let actionContext = await ActionSdkHelper.getActionContext();
    actionContext.success && setContext(actionContext.context as actionSDK.ActionSdkContext);
});

orchestrator(initialize, () => {
    fetchCurrentContext();
    let response = Localizer.initialize();
    response ? setAppInitialized(ProgressState.Completed) : setAppInitialized(ProgressState.Failed);
});

orchestrator(callActionInstanceCreationAPI, async () => {
    var actionInstance: actionSDK.Action = {
        displayName: "Poll",
        expiryTime: getStore().settings.dueDate,
        dataTables: [
            {
                name: "",
                dataColumns: [],
                attachments: [],
            },
        ],
    };

    //create poll question
    updateTitle(getStore().title.trim());

    var pollQuestion: actionSDK.ActionDataColumn = {
        name: "0",
        valueType: actionSDK.ActionDataColumnValueType.SingleOption,
        displayName: getStore().title,
    };
    actionInstance.dataTables[0].dataColumns.push(pollQuestion);
    actionInstance.dataTables[0].dataColumns[0].options = [];

    // Create poll options
    for (var index = 0; index < getStore().options.length; index++) {
        updateChoiceText(index, getStore().options[index].trim());

        var pollChoice: actionSDK.ActionDataColumnOption = {
            name: `${index}`,
            displayName: getStore().options[index],
        };
        actionInstance.dataTables[0].dataColumns[0].options.push(pollChoice);
    }

    // Set poll responses visibility
    if (getStore().settings.resultVisibility === actionSDK.Visibility.Sender) {
        actionInstance.dataTables[0].rowsVisibility = actionSDK.Visibility.Sender;
    } else {
        actionInstance.dataTables[0].rowsVisibility = actionSDK.Visibility.All;
    }

    if (validateActionInstance(actionInstance)) {
        setSendingFlag();
        prepareActionInstance(actionInstance, toJS(getStore().context));
        await ActionSdkHelper.createActionInstance(actionInstance);
    } else {
        shouldValidateUI(true);
    }
});

function prepareActionInstance(
    actionInstance: actionSDK.Action,
    actionContext: actionSDK.ActionSdkContext
) {
    if (Utils.isEmptyString(actionInstance.id)) {
        actionInstance.id = Utils.generateGUID();
        actionInstance.createTime = Date.now();
    }
    if (Utils.isEmptyObject(actionInstance.subscriptions)) {
    }
    actionInstance.updateTime = Date.now();
    actionInstance.creatorId = actionContext.userId;
    actionInstance.actionPackageId = actionContext.actionPackageId;
    actionInstance.version = actionInstance.version || 1;
    actionInstance.dataTables[0].rowsEditable =
        actionInstance.dataTables[0].rowsEditable || true;
    actionInstance.dataTables[0].canUserAddMultipleRows =
        actionInstance.dataTables[0].canUserAddMultipleRows || false;
    actionInstance.dataTables[0].rowsVisibility =
        actionInstance.dataTables[0].rowsVisibility || actionSDK.Visibility.All;

    let isPropertyExists: boolean = false;

    if (actionInstance.customProperties && actionInstance.customProperties.length > 0) {
        for (let property of actionInstance.customProperties) {
            if (property.name == "Locale") {
                isPropertyExists = true;
            }
        }
    }

    if (!isPropertyExists) {
        actionInstance.customProperties = actionInstance.customProperties || [];
        actionInstance.customProperties.push({
            name: "Locale",
            valueType: actionSDK.ActionPropertyValueType.Text,
            value: actionContext.locale,
        });
    }
}