import axios from 'axios';

interface IAPIAuthorization {
    access_token: string;
    token_type: 'bearer' | string;
    expires_in: number;
    scope: string;
}

export const authenticate = async (serverUrl: string, clientId: string, clientSecret: string, scope: string): Promise<IAPIAuthorization> => {

    try {
        const params = new URLSearchParams();
        params.append('grant_type', 'client_credentials');
        params.append('client_id', clientId);
        params.append('client_secret', clientSecret);

        const authResponse = await axios({
            method: 'POST',
            url: serverUrl,
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded',
                'Accept': 'application/json'
            },
            data: params
        });
        return authResponse?.data;
    } catch (error) {
        console.log("Could not obtain oauth token from sNow, aborting.");
        throw error;
    }
};

