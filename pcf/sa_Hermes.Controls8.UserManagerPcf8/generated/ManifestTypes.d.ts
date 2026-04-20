/*
*This is auto generated from the ControlManifest.Input.xml file
*/

// Define IInputs and IOutputs Type. They should match with ControlManifest.
export interface IInputs {
    userAccountColumn: ComponentFramework.PropertyTypes.StringProperty;
    typeUserColumn: ComponentFramework.PropertyTypes.StringProperty;
    titleUserColumn: ComponentFramework.PropertyTypes.StringProperty;
    roleColumn: ComponentFramework.PropertyTypes.StringProperty;
    lieuxColumn: ComponentFramework.PropertyTypes.StringProperty;
    lieuxTitleColumn: ComponentFramework.PropertyTypes.StringProperty;
    createdOnColumn: ComponentFramework.PropertyTypes.StringProperty;
    usersJson: ComponentFramework.PropertyTypes.StringProperty;
    typeUserChoices: ComponentFramework.PropertyTypes.StringProperty;
    titleUserChoices: ComponentFramework.PropertyTypes.StringProperty;
    roleChoices: ComponentFramework.PropertyTypes.StringProperty;
    users: ComponentFramework.PropertyTypes.DataSet;
}
export interface IOutputs {
    requestedAction?: string;
    selectedIds?: string;
    formData?: string;
    eventToken?: string;
    saveSignal?: string;
    searchTerm?: string;
    selectedId?: string;
}