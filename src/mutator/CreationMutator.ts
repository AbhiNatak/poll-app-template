import { mutator } from 'satcheljs';
import { setContext, setSendingFlag, goToPage, updateTitle, setAppInitialized, updateSettings, updateChoiceText, deleteChoice, shouldValidateUI, addChoice } from './../actions/CreationActions';
import * as actionSDK from "@microsoft/m365-action-sdk";
import { Utils } from '../utils/Utils';
import getStore from "../store/CreationStore";

mutator(setContext, (msg) => {
    const store = getStore();
    store.context = msg.context;
    store.initPending = false;
    if (!Utils.isEmptyObject(store.context.viewData)) {
        const viewData = JSON.parse(store.context.viewData);
        const actionInstance: actionSDK.Action = viewData["actionInstance"];
        getStore().title =
            actionInstance.dataTables[0].dataColumns[0].displayName;
        let options = actionInstance.dataTables[0].dataColumns[0].options;
        // clearing the options since it is always initialize with 2 empty options.
        getStore().options = [];
        options.forEach((option) => {
            getStore().options.push(option.displayName);
        });

        if (
            actionInstance.dataTables[0].rowsVisibility ===
            actionSDK.Visibility.Sender
        ) {
            getStore().settings.resultVisibility = actionSDK.Visibility.Sender;
        } else {
            getStore().settings.resultVisibility = actionSDK.Visibility.All;
        }
        getStore().settings.dueDate = actionInstance.expiryTime;
    }
});


mutator(setSendingFlag, () => {
    const store = getStore();
    store.sendingAction = true;
});

mutator(goToPage, (msg) => {
    const store = getStore();
    store.currentPage = msg.page;
});



mutator(addChoice, () => {
    const store = getStore();
    const optionsCopy = [...store.options];
    optionsCopy.push("");
    store.options = optionsCopy;
});

mutator(shouldValidateUI, (msg) => {
    const store = getStore();
    store.shouldValidate = msg.shouldValidate;
});

mutator(deleteChoice, (msg) => {
    const store = getStore();
    const optionsCopy = [...store.options];
    optionsCopy.splice(msg.index, 1);
    store.options = optionsCopy;
});

mutator(updateChoiceText, (msg) => {
    const store = getStore();
    const optionsCopy = [...store.options];
    optionsCopy[msg.index] = msg.text;
    store.options = optionsCopy;
});

mutator(updateTitle, (msg) => {
    const store = getStore();
    store.title = msg.title;
});

mutator(updateSettings, (msg) => {
    const store = getStore();
    store.settings = msg.settingProps;
});

mutator(setAppInitialized, (msg) => {
    const store = getStore();
    store.isInitialized = msg.state;
});