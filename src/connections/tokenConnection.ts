import { IConnectionSchema } from "@cognigy/extension-tools";

export const tokenConnection: IConnectionSchema = {
        type: "token",
        label: "token",
        fields: [
            { fieldName: "token" },
        ]
};