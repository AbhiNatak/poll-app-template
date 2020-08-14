import { createStore } from "satcheljs";
import * as actionSDK from "@microsoft/m365-action-sdk";
import { Utils } from "../utils/Utils";
import { ISettingsComponentProps } from "./../components/Creation/Settings";
import { ProgressState } from "./../utils/SharedEnum";
import './../orchestrators/CreationOrchestrator';
import './../mutator/CreationMutator';

export enum Page {
    Main,
    Settings,
}

interface IPollCreationStore {
    context: actionSDK.ActionSdkContext;
    initPending: boolean;
    title: string;
    maxOptions: number;
    options: string[];
    settings: ISettingsComponentProps;
    shouldValidate: boolean;
    sendingAction: boolean;
    isInitialized: ProgressState;
    currentPage: Page;
}

const store: IPollCreationStore = {
    context: null,
    initPending: true,
    title: "",
    maxOptions: 10,
    options: ["", ""],
    settings: {
        resultVisibility: actionSDK.Visibility.All,
        dueDate: Utils.getDefaultExpiry(7).getTime(),
        strings: null,
    },
    shouldValidate: false,
    sendingAction: false,
    currentPage: Page.Main,
    isInitialized: ProgressState.NotStarted,
};

export default createStore<IPollCreationStore>("cerationStore", store);
