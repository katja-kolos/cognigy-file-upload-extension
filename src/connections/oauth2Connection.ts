import { IConnectionSchema } from "@cognigy/extension-tools";

export const OAuth2Connection: IConnectionSchema = {
        type: "http_oauth2",
        label: "OAuth2 for HTTP Requests",
        fields: [
            // {fieldName: "myOauthConnection"},
            { fieldName: "oAuth2Url" },
            { fieldName: "oAuth2ClientId" },
            { fieldName: "oAuth2ClientSecret" },
            { fieldName: "oAuth2Scope" },
        ]
};