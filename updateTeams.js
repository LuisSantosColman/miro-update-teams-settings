/* 
DISCLAIMER:
The content of this project is subject to the Miro Developer Terms of Use: https://miro.com/legal/developer-terms-of-use/
This script is provided only as an example to illustrate how to identify Miro Teams with no Boards within and to remove these empty Teams.
The usage of this script is at the sole discretion and responsibility of the customer and is to be tested thoroughly before running it on Production environments.

Script Author: Luis Colman (luis.s@miro.com) | Global Senior Solutions Architect at Miro | LinkedIn: https://www.linkedin.com/in/luiscolman/
*/

const IS_TEST = true; // Change to false to perform team deletions
const TOKEN = 'YOUR_MIRO_REST_API_TOKEN'; // Replace with your Miro REST API token
const MIRO_ORGANIZATION_ID = 'YOUR_MIRO_ORGANIZATION_ID'; // Replace with your Miro Company ID
/* 
* The variable TEAM_SETTINGS_PAYLOAD is the JSON payload of the settings you want to change. Modify this payload to match the specific settings you would like to change
* To see examples of payloads for the existing Miro settings, please see this mapping table: https://docs.google.com/spreadsheets/d/1pd9WuR_7XWVg84h8c7I3kxaYmYnvNKlAcT5f9M5hvUg/edit?usp=sharing
* Official Team Settings API documentation: https://developers.miro.com/reference/enterprise-update-team-settings
* Adjust the below JSON file with the respective paylod to update the settings you want to update
*/
const TEAM_SETTINGS_PAYLOAD = {
    'teamInvitationSettings': {
        'inviteExternalUsersEnabled': true, // change to false if Guests should not be allowed in the Team
        'inviteExternalUsers': 'allowed' // change to not_allowed if Guests should not be allowed in the Team
    },
    'teamAccountDiscoverySettings': {
        'accountDiscovery': 'request'
    }
};

/* SCRIPT BEGIN */

/* Variables - BEGIN */
const fs = require('fs');
let getUserErrors = {};
let userObject = {};
let teams = {};
let getIndividualTeamsErrors = {};
let errorRetryCount = 0;
let numberOfRequests = 520;
let numberOfRequestsForUpdate = 52;
let affectedTeams = {};
let results = {};
let getBoardsErrors = {};
let teamsToRemove = {};
/* Variables - END */

/* Functions - BEGIN */

/* Function to get the value of a query parameter of a string URL */
function getParameterByName(name, url) {
    if (!url) url = window.location.href;
    name = name.replace(/[\[\]]/g, "\\$&");
    const regex = new RegExp("[?&]" + name + "(=([^&#]*)|&|#|$)"),
        results = regex.exec(url);
    if (!results) return null;
    if (!results[2]) return "";
    //return decodeURIComponent(results[2].replace(/\+/g, " "));
    return results[2];
}

/* Functions to hold script execution (to allow replenishing credits when hitting the SCIM API rate limits) */
const delay = ms => new Promise(res => setTimeout(res, ms));
const holdScriptExecution = async (ms) => {
    console.log('**** Rate limit hit - Delaying execution for ' + (ms / 1000) + ' seconds to replenish rate limit credits - Current time: ' + new Date() + ' ***');
    
    let elapsedSeconds = 0;
    const intervalId = setInterval(() => {
        elapsedSeconds++;
        console.log(`${elapsedSeconds} second(s) passed...`);
    }, 1000);

    await delay(ms);
    
    clearInterval(intervalId); // Stop the interval when delay is over
    console.log('**** Resuming script - Current time: ' + new Date() + ' ***');
};

/* Convert JSON to CSV */
function jsonToCsv(jsonData) {
    if (jsonData) {
        let csv = '';
        let headers;
        
        // Get the headers
        if (IS_TEST) { 
            headers = (Object?.keys(jsonData).length > 0 ? Object?.keys(jsonData[Object?.keys(jsonData)[0]]) : ['NO DELETIONS OCCURRED - TEST MODE WAS ON']);
        }
        else {
            headers = (Object?.keys(jsonData).length > 0 ? Object?.keys(jsonData[Object?.keys(jsonData)[0]]) : ['NO DELETIONS OCCURRED - NO DATA TO SHOW']);
        }
        csv += headers.join(',') + '\n';
        
        // Helper function to escape CSV special characters
        const escapeCSV = (value) => {
            if (Array.isArray(value)) {
                // Join array values with a comma followed by a space
                value = value.join(', ');
            }
            if (typeof value === 'string') {
                // Escape double quotes
                if (value.includes('"')) {
                    value = value.replace(/"/g, '""');
                }
            }
            // Wrap the value in double quotes to handle special CSV characters
            value = `"${value}"`;
            return value;
        };
    
        // Add the data
        Object.keys(jsonData).forEach(function(row) {
            let data = headers.map(header => escapeCSV(jsonData[row][header])).join(',');
            csv += data + '\n';
        });

        return csv;
    }
}

/* Function to create reports */
function addReportsForNodeJS() {
    let content;
    let directory = 'miro_teams_update_output_files';
    if (!fs.existsSync(directory)) {
        fs.mkdirSync(directory);
    }

    content = JSON.stringify(teams, null, '2');
    filePath = 'miro_teams_update_output_files/Miro_Teams_Overview.json';
    fs.writeFileSync(filePath, content);

    content = jsonToCsv(teams);
    filePath = 'miro_teams_update_output_files/Miro_Teams_Overview.csv';
    fs.writeFileSync(filePath, content);

    content = JSON.stringify(results, null, '2');
    filePath = 'miro_teams_update_output_files/Miro_Teams_Update_Results.json';
    fs.writeFileSync(filePath, content);

    content = jsonToCsv(results);
    filePath = 'miro_teams_update_output_files/Miro_Teams_Update_Results.csv';
    fs.writeFileSync(filePath, content);

    if (Object.keys(getIndividualTeamsErrors).length > 0) {
        content = JSON.stringify(getIndividualTeamsErrors, null, '2');
        filePath = 'miro_teams_update_output_files/Script_Errors.json';
        fs.writeFileSync(filePath, content);

        content = jsonToCsv(getIndividualTeamsErrors);
        filePath = 'miro_teams_update_output_files/Script_Errors.csv';
        fs.writeFileSync(filePath, content);
    }
}

/* Function to call Miro API teams */
async function callAPI(url, options) {
    async function manageErrors(response) {
        if(!response.ok){
            var parsedResponse = await response.json();
            var responseError = {
                status: response.status,
                statusText: response.statusText,
                requestUrl: response.url,
                errorDetails: parsedResponse
            };
            throw(responseError);
        }
        return response;
    }

    var response = await fetch(url, options)
    .then(manageErrors)
    .then((res) => {
        if (res.ok) {
            var rateLimitRemaining = res.headers.get('X-RateLimit-Remaining');
            return res[res.status == 204 ? 'text' : 'json']().then((data) => ({ status: res.status, rate_limit_remaining: rateLimitRemaining, body: data }));
        }
    })
    .catch((error) => {
        console.error('Error:', error);
        return error;
    });
    return response;
}

/* Function to update teams with no boards within */
async function updateTeams(numberOfRequestsForUpdates) {
    let totalItems;
    let batchUrls;

    let reqHeaders = {
        'cache-control': 'no-cache, no-store',
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + TOKEN
    };

    let payload = JSON.stringify(TEAM_SETTINGS_PAYLOAD);

    let reqGetOptions = {
        method: 'PATCH',
        headers: reqHeaders,
        body: payload
    };

    totalItems = Object.keys(teams);
    getRemainingTeams = {};

    for(let i=0; i < totalItems.length; i++) {
        getRemainingTeams[totalItems[i]] = { team_name: teams[totalItems[i]].team_name, team_id: teams[totalItems[i]].team_id }
    }

    getProcessedTeams = {};
    let processedUrls = [];
    let batchSize;

    while (Object.keys(getRemainingTeams).length > 0) {
        var apiUrl = `https://api.miro.com/v2/orgs/${MIRO_ORGANIZATION_ID}/teams`;
        
        // Calculate the number of items remaining to fetch
        const remainingItems = totalItems.length - (Object.keys(getProcessedTeams).length);

        if (Object.keys(getIndividualTeamsErrors).length === 0) {
            // Calculate the number of calls to make in this batch
            batchSize = Math.min(numberOfRequestsForUpdates, Math.ceil(remainingItems / 1));
            batchUrls = Array.from({ length: batchSize }, (_, index) => `${apiUrl}/${Object.keys(getRemainingTeams)[index]}/settings`);
        }
        else {
            console.log('Errors found - retrying failed requests');
            await holdScriptExecution(43000); 
            batchSize = Object.keys(getIndividualTeamsErrors).length;
            batchUrls = Array.from({ length: batchSize }, (_, index) => `${Object.keys(getIndividualTeamsErrors)[index]}`);
            processedUrls.forEach(function(item) {
                let urlIndex = batchUrls.indexOf(item);
                if (urlIndex !== -1) {
                    batchUrls.splice(urlIndex, 1);
                }
            });
            errorRetryCount = errorRetryCount + 1;
            console.log(`errorRetryCount --> ${errorRetryCount}`);
            if (errorRetryCount < 8) {
                if (errorRetryCount === 7) { 
                    console.log('This is the 7th and last attempt to retry failed "updateTeams" calls...');
                }
            }
            else {
                console.log('Maximum amount of retry attempts for failed "updateTeams" calls reached (7). Please review the "getIndividualTeamsErrors" object to find out what the problem is...');
                return false;
            }
        }
        if (Object.keys(getIndividualTeamsErrors).length > 0) { 
            console.log(`Failed API calls to retry below: -----`); 
        }

        if (batchUrls.length > 0) {

            console.log(`.........API URLs in this the batch are:`);
            console.table(batchUrls);

            try {       
                const promisesWithUrls = batchUrls.map(url => {
                    const promise = fetch(url, reqGetOptions)
                        .catch(error => {
                            // Check if the error is a response error
                            if (error instanceof Response) {
                                // Capture the HTTP error code and throw it as an error
                                let teamId = url.split('/');
                                teamId = teamId[7];
                                if (!getIndividualTeamsErrors[url]) {
                                    getIndividualTeamsErrors[url] = { team: teamId, url: url, error: error.status, errorMessage: error.statusText };
                                }
                                console.error({ team: teamId, url: url, errorMessage: errorMessage });
                                return Promise.reject(error);
                            } else {
                                // For other types of errors, handle them as usual
                                throw error;
                            }
                        });
                    return { promise, url };
                });

                // Fetch data for each URL in the batch
                const batchResponses = await Promise.allSettled(promisesWithUrls.map(({ promise }) => promise));
                for (let i = 0; i < batchResponses.length; i++) {
                    let { status, value, reason } = batchResponses[i];
                    if (status === 'fulfilled') {
                        let teamId = value.url.split('/');
                        teamId = teamId[7];
                        if (value.ok) {
                            errorRetryCount = 0;
                            if (processedUrls.indexOf(value.url) === -1) {
                                results[teamId] = {
                                    team_id: teamId,
                                    team_name: teams[teamId].team_name,
                                    result: `Team ${teamId} successfully updated`
                                };
                                processedUrls.push(value.url);
                                delete getRemainingTeams[teamId];
                                if (!getProcessedTeams[teamId]) {
                                    getProcessedTeams[teamId] = { team_id: teamId, team_name: teams[teamId].team_name };
                                }
                                if (getIndividualTeamsErrors[value.url]) {
                                    delete getIndividualTeamsErrors[value.url];
                                }
                                console.log(`Team ${teamId} successfully updated - Team ${Object.keys(getProcessedTeams).length} out of ${totalItems.length}`);
                            }
                        }
                        else if (value.status === 429) {
                            if (!getIndividualTeamsErrors[value.url]) {
                                getIndividualTeamsErrors[value.url] = { team_id: teamId, team_name: teams[teamId].team_name, errorCode: value.status, errorMessage: 'Rate limit reached' };
                            }
                        }
                        else if (value.status === 500) {
                            if (processedUrls.indexOf(value.url) === -1) {
                                if (!getIndividualTeamsErrors[value.url]) {
                                    getIndividualTeamsErrors[value.url] = { team_id: teamId, team_name: teams[teamId].team_name, errorCode: value.status, errorMessage: reason };
                                }
                                console.log(`Team ${teamId} returned a 500 Error - Call will be re-tried - Processed teams: ${Object.keys(getProcessedTeams).length} out of ${totalItems.length}`);
                            }
                        }
                        else {
                            let batchData = await value.json();
                            if (!getIndividualTeamsErrors[value.url]) {
                                getIndividualTeamsErrors[value.url] = { team_id: teamId, team_name: teams[teamId].team_name, errorCode: value.status, errorMessage: batchData.message };
                            }
                            console.log(`Error - Could not update Team - Processed teams: ${Object.keys(getProcessedTeams).length} out of ${totalItems.length} - Current Team: ${teamId}`);
                            console.dir(batchData);
                        }
                    }
                    else {
                        let index = batchResponses.indexOf({ status, value, reason });
                        let failedUrl = promisesWithUrls[index].url;
                        let teamId = failedUrl.split('/');
                        teamId = teamId[7];
                        if (!getIndividualTeamsErrors[failedUrl]) {
                            getIndividualTeamsErrors[failedUrl] = { team: teamId, url: failedUrl, error: status, errorMessage: value.statusText };
                        }
                        console.error(`Custom Error - API URL --> ${failedUrl}:`, reason);
                    }
                }
            }
            catch (error) {
                console.error(error);
            }
        }
    }
    return true;
}

/* Function to get all Teams in the Miro account */
async function getTeams(orgId, cursor) {
    let reqHeaders = {
        'cache-control': 'no-cache, no-store',
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + TOKEN
    };

    let reqGetOptions = {
        method: 'GET',
        headers: reqHeaders,
        body: null
    };

    let url = `https://api.miro.com/v2/orgs/${orgId}/teams` + (cursor ? `?cursor=${cursor}` : '');
    console.log('Getting Miro Teams - API URL --> : ' + url);
    let listTeams = await callAPI(url, reqGetOptions);
    if (listTeams.status === 200) {
        for(let i=0; i < listTeams.body.data.length; i++) {
            let teamId = listTeams.body.data[i].id;
            teams[teamId] = {
                miro_org_id: orgId,
                team_name: listTeams.body.data[i].name.toString(),
                team_id: teamId.toString()
            };
        }

        if (listTeams.body.cursor) {
            await getTeams(orgId, listTeams.body.cursor);
        }
        else {
            console.log('Getting Miro Teams COMPLETE...');
           // await getBoards(numberOfRequests);

            if (Object.keys(getIndividualTeamsErrors).length === 0) {

                if (!IS_TEST) {
                    if (Object.keys(teams).length > 0) {
                        console.log('Preparing to update Team Settings...');
                        await updateTeams(numberOfRequestsForUpdate);
                        if (Object.keys(getIndividualTeamsErrors).length === 0) {
                            console.log('Updating Team Settings COMPLETE...');
                        }
                    }
                }
                else {
                    console.log('TEST MODE ON: Skipping Team Settings changes...');
                }
            }
            
            addReportsForNodeJS();
            console.log(`Script end time: ${new Date()}`);

            console.log('\n======================================');
            console.log('IMPORTANT: Please review script results within the folder "miro_teams_update_output_files" in the directory of this script...');
            console.log('========================================\n');
            
            console.log('********** END OF SCRIPT **********\n\n');
            return true;
        }
    }
    else {
        if (!getIndividualTeamsErrors[url]) {
            getIndividualTeamsErrors[url] = { errorCode: listTeams.status, errorMessage: listTeams?.body?.message };
            console.error(listTeams);
            addReportsForNodeJS();
            return listTeams;
        }
    }
    if (listTeams.rate_limit_remaining === '0') {
        await holdScriptExecution(61000);
    }
}

getTeams(MIRO_ORGANIZATION_ID);
