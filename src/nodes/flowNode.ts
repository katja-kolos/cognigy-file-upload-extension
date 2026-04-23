import { createNodeDescriptor, INodeFunctionBaseParams } from "@cognigy/extension-tools";
import axios from "axios";
import { authenticate } from "../helper/authenticate";
import addToStorage from "../helper/addToStorage";

export interface IMoveBlobParams extends INodeFunctionBaseParams {
  config: {
    // connection: any;
    // connectionSection: any;
    msToken: string;
    authentication: "none" | "token" | "oauth2";
    connection: {
      oAuth2Url?: string;
      oAuth2ClientId?: string;
      oAuth2ClientSecret?: string;
      oAuth2Scope?: string;
    };
    sNowConnection: any;
    // sourceUrl: string;
    // inputData: any;
    destinationUrl: string;
    storeLocation: any;
    inputKey: string;
    contextKey: string;
  };
}
export const moveBlob = createNodeDescriptor({
  type: "moveBlob", // type of your node. It needs to be a unique string in your Extension.
  defaultLabel: "Upload Teams - ServiceNow", // default values for your Node. If a new Flow Node gets created, it will get these defaults. You must set the label and the config object.
  // preview: { key: "sourceUrl", type: "text" },
  preview: { key: "destinationUrl", type: "text" },
  fields: [
    // {
    //   key : "msToken",
		// 	label: "Microsoft Authentication Token",
    //   description: "Token obtained with HTTP node necessary for inline upload from MS Teams.",
		// 	type: "cognigyText",
		// 	params: {
		// 	  required: true
		// 	}
		// },
    {
      key : "authentication",
			label: "Authentication Type",
      description: "Whether to use authentication and if yes, which kind of authentication",
			type: "select",
			defaultValue: "oauth2",
			params: {
			  required: true,
				options: [
          {
            label: "None",
            value: "none"
          },
					{
						label: "Existing ServiceNow Token",
						value: "token"
					},
					{
						label: "OAuth2",
						value: "oauth2"
					},
				],
			}
		},
    {
      key: "connection",
      label: "OAuth2 for HTTP Requests",
      type: "connection",
      params: {
          connectionType: "http_oauth2",
          label: "OAuth2 Parameters",
          required: true,
      },
      condition: {
        key: "authentication",
        value: "oauth2",
      }
    },
    {
      key: "sNowConnection",
      label: "Existing authentication token",
      description: "Service Now OAuth2 access token obtained previously",
      type: "cognigyText",
      params: { required: true },
      condition: {
        key: "authentication",
        value: "token",
      }
    },
    // {
    //   key: "inputData",
    //   label: "Input data (attachments)",
    //   description: "Pass {{input.data.request}}",
    //   type: "json",
    //   params: { required: true }
    // },
    // {
    //   key: "sourceUrl",
    //   label: "Source URL (MS Teams > NiCE Sharepoint)",
    //   description: "Source tmp URL. Format: https://niceonline-my.sharepoint.com/personal/user_name_nice_com/_layouts/15/download.aspx?UniqueId=unique_id&Translate=false&tempauth=v1.api_key&ApiVersion=2.0",
    //   type: "cognigyText",
    //   params: { required: true }
    // },
    {
      key: "destinationUrl",
      label: "Destination Blob URL in ServiceNow",
      description: "Destination blob URL, e.g. https://myaccount.blob.core.windows.net/container-name/renamed.txt",
      type: "cognigyText",
      params: { required: true }
    },
    {
      key: 'storeLocation',
      type: 'select',
      label: 'Where to store the result',
      defaultValue: 'input',
      params: {
        options: [
          {
            label: 'Input',
            value: 'input'
          },
          {
            label: 'Context',
            value: 'context'
          }
        ],
        required: true
      }
    },
    {
      key: 'inputKey',
      type: 'cognigyText',
      label: 'Input Key to store Result',
      defaultValue: 'case',
      condition: {
        key: 'storeLocation',
        value: 'input'
      }
    },
    {
      key: 'contextKey',
      type: 'cognigyText',
      label: 'Context Key to store Result',
      defaultValue: 'case',
      condition: {
        key: 'storeLocation',
        value: 'context'
      }
    } // user interface that will be generated for your node-config. You have to add a field definition per key in your config object. You reference the key in your config using the key property in the field definition. The label is used in the UI as well. The type gives you a variety of possibilities - you can, e.g., say that your field should be of type cognigyText or, e.g., of type JSON or toggle
  ],
  sections: [
    {
      key: "connectionSection",
			label: "Authentication",
			defaultCollapsed: false,
			fields: [
        "authentication",
        "connection",
				"sNowConnection",
      ]
		},
    {
      key: 'storage',
      label: 'Storage Option',
      defaultCollapsed: true,
      fields: [
        'storeLocation',
        'inputKey',
        'contextKey'
      ]
    }
  ],
  form: [
    // { type: "field", key: "msToken" },
    { type: "section", key: "connectionSection" },
    // { type: "field", key: "inputData" },
    // { type: "field", key: "sourceUrl" },
    { type: "field", key: "destinationUrl" },
    { type: "section", key: "storage" },
  ],
  appearance: { color: "#181818" },
  function: async ({ cognigy, config }: IMoveBlobParams) => {
    const { api } = cognigy;

    // const { msToken, authentication, connection, sNowConnection, inputData, destinationUrl, storeLocation, contextKey, inputKey } = config;
    const {
      // msToken,
      authentication,
      connection,
      sNowConnection,
      // inputData,
      destinationUrl,
      storeLocation,
      contextKey,
      inputKey
    } = config;
    const { oAuth2Url, oAuth2ClientId, oAuth2ClientSecret, oAuth2Scope } = connection as any;
    const msToken = api.getContext("botFrameworkTokenResponse")?.result?.access_token;
    if (!msToken) {
      api.log("error", "Microsoft token not found in context. Make sure to pass it into the context with key 'botFrameworkTokenResponse.result.access_token'.");
      return;
    }
    try {
      api.log('info', `Will move file from Teams to ServiceNow`);
      let bearerToken: string | undefined;
      if (authentication === "token") {
        bearerToken = sNowConnection;
        api.log("info", "Using existing ServiceNow token");
      } else if (authentication === "oauth2") {
        api.log("info", "Requesting ServiceNow access token...");
        const oauthResponse = await authenticate(
          oAuth2Url,
          oAuth2ClientId,
          oAuth2ClientSecret,
          oAuth2Scope
        );
        bearerToken = oauthResponse?.access_token;
      }
      if (!bearerToken) {
        api.log("error", "Bearer token is undefined. Aborting file upload.");
        return;
      }
      // a) extract sourceURL from input, instead of giving it in the flow
      // (or make it possible to modify in the flow but add option extract from input)
      // b) construct destination URL based on input and context
      // (or make it possible to override in the flow, but add option to construct based on input and context)
      let sourceUrl;
      let sourceResponse;
      // const attachments = Array.isArray(inputData)
      //   ? inputData
      //   : inputData?.attachments;
      const attachments = api.getContext("attachments");
      const tableName = api.getContext("tablename");
      const sysId = api.getContext("snowResponse").result?.result[0]?.sys_id;
      api.log("info", `Extracted table name: ${tableName} and sys_id: ${sysId} from context`);

      if (!Array.isArray(attachments) || attachments.length === 0) {
        api.log("error", "No attachments found. Pass {{input.data.request.attachments}} or an object with .attachments");
        return;
      }
      api.log("info", `Found ${attachments.length - 1} attachment(s)`); // 1 attachment is the text/html content which we want to skip
      const uploadResults = [];
      for (const attachment of attachments) {
        if (attachment.contentType === "text/html") {
          api.log("info", "Skipping attachment with content type text/html");
          continue;
        }
        // 1) Download file bytes from the source URL, depending on the type of upload (inline vs regular) - contentUrl vs downloadUrl
        if (attachment.downloadUrl) {
          sourceUrl = attachment.downloadUrl;
          api.log('info', `Regular upload. Downloading file from ${sourceUrl}`);
          sourceResponse = await axios.get(sourceUrl, {
            responseType: 'arraybuffer',
            maxContentLength: Infinity,
            maxBodyLength: Infinity
          });
          const contentType = sourceResponse.headers['content-type'] || 'application/octet-stream';
          const fileBytes = sourceResponse.data;
          const size = fileBytes?.byteLength ?? fileBytes?.length ?? 0;
          api.log('info', `Downloaded ${size} bytes (content-type: ${contentType})`);

          // 2) Upload the bytes to the destination URL
          api.log('info', `Uploading file to ${destinationUrl}`);
          const uploadResponse = await axios({
            method: 'post',
            url: destinationUrl,
            headers: {
              'Authorization': `Bearer ${bearerToken}`,
              'Accept': 'application/json',
              'Content-Type': contentType // Reusing the content-type from the source
            },
            data: fileBytes, // The Buffer/ArrayBuffer from the GET request
            // ?table_name={{context.tablename}}&table_sys_id={{context.sys_id}}&file_name={{input.fileName}}
            params: {
              table_name: tableName,
              table_sys_id: sysId,
              file_name: attachment.name || `uploaded_file_${Date.now()}`
            },
            maxContentLength: Infinity,
            maxBodyLength: Infinity
          });
          api.log('info', `Successfully uploaded attachment: ${uploadResponse.data.result.sys_id}`);
          uploadResults.push(uploadResponse.data);
        } else if (attachment.content?.downloadUrl) {
          sourceUrl = attachment.content.downloadUrl;
          api.log('info', `Regular upload. Downloading file from ${sourceUrl}`);
          sourceResponse = await axios.get(sourceUrl, {
            responseType: 'arraybuffer',
            maxContentLength: Infinity,
            maxBodyLength: Infinity
          });
          const contentType = sourceResponse.headers['content-type'] || 'application/octet-stream';
          const fileBytes = sourceResponse.data;
          const size = fileBytes?.byteLength ?? fileBytes?.length ?? 0;
          api.log('info', `Downloaded ${size} bytes (content-type: ${contentType})`);

          // 2) Upload the bytes to the destination URL
          api.log('info', `Uploading file to ${destinationUrl}`);
          const uploadResponse = await axios({
            method: 'post',
            url: destinationUrl,
            headers: {
              'Authorization': `Bearer ${bearerToken}`,
              'Accept': 'application/json',
              'Content-Type': contentType // Reusing the content-type from the source
            },
            data: fileBytes, // The Buffer/ArrayBuffer from the GET request
            // ?table_name={{context.tablename}}&table_sys_id={{context.sys_id}}&file_name={{input.fileName}}
            params: {
              table_name: tableName,
              table_sys_id: sysId,
              file_name: attachment.name || `uploaded_file_${Date.now()}`
            },
            maxContentLength: Infinity,
            maxBodyLength: Infinity
          });
          api.log('info', `Successfully uploaded attachment: ${uploadResponse.data.result.sys_id}`);
          uploadResults.push(uploadResponse.data);
        } else if (attachment.contentUrl) {
          // this will only be images
          // but API returns application/octet-stream
          // when upload, pass image/png (?)
          sourceUrl = attachment.contentUrl;
          api.log('info', `Inline upload. Downloading file from ${sourceUrl}`);
          sourceResponse = await axios.get(sourceUrl, {
            headers: {  'Authorization': `Bearer ${msToken}` }, // Use the Microsoft token for inline uploads
            responseType: 'arraybuffer',
            maxContentLength: Infinity,
            maxBodyLength: Infinity
          });
          const contentType = 'image/png';
          const fileBytes = sourceResponse.data;
          const size = fileBytes?.byteLength ?? fileBytes?.length ?? 0;
          api.log('info', `Downloaded ${size} bytes (content-type: ${contentType})`);

          // 2) Upload the bytes to the destination URL
          api.log('info', `Uploading file to ${destinationUrl}`);
          const uploadResponse = await axios({
            method: 'post',
            url: destinationUrl,
            headers: {
              'Authorization': `Bearer ${bearerToken}`,
              'Accept': 'application/json',
              'Content-Type': contentType // Reusing the content-type from the source
            },
            // ?table_name={{context.tablename}}&table_sys_id={{context.sys_id}}&file_name={{input.fileName}}
            params: {
              table_name: tableName,
              table_sys_id: sysId,
              file_name: attachment.name || `uploaded_file_${Date.now()}`
            },
            data: fileBytes, // The Buffer/ArrayBuffer from the GET request
            maxContentLength: Infinity,
            maxBodyLength: Infinity
          });
          api.log('info', `Successfully uploaded attachment: ${uploadResponse.data.result.sys_id}`);
          uploadResults.push(uploadResponse.data);
        }
      }
    addToStorage({
      api,
      storeLocation,
      contextKey,
      inputKey,
      data: uploadResults
    });
    } catch (error) {
      const errorMsg = error.response ? JSON.stringify(error.response.data) : error.message;
      api.log('error', `ServiceNow Upload Failed: ${errorMsg}`);
    }
  }
});