import { IInputs, IOutputs } from "./generated/ManifestTypes";

interface UserRecordViewModel {
    id: string;
    datasetId: string;
    userAccount: string;
    typeUser: string;
    titleUser: string;
    role: string;
    lieux: string[];
    createdOn?: Date;
}

type SortDir = "asc" | "desc";
type SortKey = "userAccount" | "typeUser" | "titleUser" | "role" | "createdOn" | "lieux";

interface ColFilter {
    userAccount: string;
    typeUser: string;
    titleUser: string;
    role: string;
    lieux: string;
}

export class UserManagerPcf8 implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    // ...existing code copied from UserManagePcf8...
}
