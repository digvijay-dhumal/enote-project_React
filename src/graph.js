import { graphConfig } from "./config";
 
/**
 * Attaches a given access token to a MS Graph API call. Returns information about the user
 * @param accessToken
 */
export async function callMsGraph(accessToken) {
    // console.log("Access Token :" , accessToken);
    const headers = new Headers();
    const bearer = `Bearer ${accessToken}`;
    headers.append("Authorization", bearer);
 
    const options = {
        method: "GET",
        headers: headers
    };
 
    try {
        const response = await fetch(graphConfig.graphMeEndpoint, options);
        if (response.ok) {
            return response.json();
        } else {
            throw new Error(`Failed to fetch data: ${response.status} - ${response.statusText}`);
        }
    } catch (error) {
        throw new Error(`Error fetching data: ${error.message}`);
    }
}
 