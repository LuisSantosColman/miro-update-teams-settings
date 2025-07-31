# Miro - Update Miro Teams Settings via REST API (Node.js)

This repository contains a script in JavaScript/Node.js to delete empty Miro teams using Miro's REST API.

The script provides a __TEST MODE__ flag to run it without performing any changes. (See __step 5__ further below).

## Disclaimer
The content of this project is subject to the Miro Developer Terms of Use: https://miro.com/legal/developer-terms-of-use/<br>
This script is provided as an example to illustrate how to update Miro team settings via API.
The usage of this script is at the sole discretion and responsibility of the customer and is to be tested thoroughly before running it on Production environments.

## Requirements

* [NodeJS 16.x or higher installed](https://nodejs.org/en/download/)
* You must be a __Company Admin__ in your Miro account, or at least the user generating the token must be a __Company Admin__ in your Miro account (see steps 3 and 4)

## Step 1. Install Node.js

1.1. If you already have Node.js installed in your local machine, you may skip this step.

1.2. If you do not have Node.js installed, proceed to download it [here](https://nodejs.org/en/download/) and proceed to install Node with the downloaded file. (Feel free to use the command line to download and install Node if preferred).

## Step 2. Create directory for your script files

2.1. In your local machine create a folder in the desired location where you will store the files within this repository.

2.2. Download this repository as .zip and extract the files within into the directory created, or clone this repository into the desired location in your local machine.

## Step 3. Create a Developer Team in Miro

3.1. If you already have a Miro Developer Team, you may skip this step.

3.2. If you do not have yet a Miro Developer Team, please visit this [Miro Help](https://help.miro.com/hc/en-us/articles/4766759572114-Enterprise-Developer-teams) page and follow the instructions within the article to create an Enterprise Developer Team for your Miro Enterprise Account.

## Step 4. Create a Miro App to get a REST API Token

4.1. To create a new application on your Miro Enterprise account using the Enterprise Developer team, navigate to __[Profile settings](https://help.miro.com/hc/en-us/articles/4408879513874-Profile-settings) > Your apps__, agree to the terms and conditions, and click on __+ Create new app__.

<img src="https://miro-org.s3.eu-central-1.amazonaws.com/board_classification/Create_new_app.png" alt="Accept app terms screenshot" width="502" />

4.2. Insert the desired app name (e.g. __Delete Empty Teams Script__), select your Developer team for the application and click on __Create app__.

<img src="https://miro-org.s3.eu-central-1.amazonaws.com/board_classification/delete-empty-teams-create-app.jpeg" alt="Create app screenshot" width="502" />

4.3. On the app page, scroll down and select the following scopes of access to grant to your REST API token:<br><br>
  `organizations:teams:read`<br>
  `organizations:teams:write`<br>

<img src="https://miro-org.s3.eu-central-1.amazonaws.com/board_classification/update_teams_scopes.png" alt="API token scopes" width="700" />

4.4. Click on __Install app and get OAuth token__

<img src="https://miro-org.s3.eu-central-1.amazonaws.com/board_classification/install_and_get_token_screenshot1.png" alt="Install and and get token screenshot" width="700" />

4.5. Select any Production team within your Enteprise account, the token will apply for the entire account based on the scopes set on step 5.3 and click on __Add__

<img src="https://miro-org.s3.eu-central-1.amazonaws.com/board_classification/select_team_screenshot.png" alt="Install and and get token screenshot" width="502" />

4.6. You will see the __REST API token__. Copy this token and store it in a secure place. You will need it when running the scripts.

<img src="https://miro-org.s3.eu-central-1.amazonaws.com/board_classification/get_access_token_screenshot.png" alt="Install and and get token screenshot" width="502" />

4.7. Find your __Miro Organization ID__ as you will need it when running the scripts. You will find your __Miro Organization ID__ in the URL of the page where you received the REST API token

<img src="https://miro-org.s3.eu-central-1.amazonaws.com/board_classification/get_miro_org_id_screenshot.png" alt="Install and and get token screenshot" width="903" />

## Step 5. Run script `updateTeams.js` using the command line (CLI)

5.1. Within the file `updateTeams.js` replace the values of the below variables at the top of the script:

  - `IS_TEST`: set to `true` to run the script in test mode (no deletions will occur only reports will be generated). Set to `false` to perform the Miro Team deletions - code line `10`
  - `TOKEN`: replace value with your REST API token from step 5.6  - code line `11`
  - `MIRO_ORGANIZATION_ID`: replace value with your Miro Organization ID from step 5.7 - code line `12`
  - `TEAM_SETTINGS_PAYLOAD`: replace value with the JSON corresponding to the team settings you would like to change - code line `19`. To locate the API payload corresponding to the Team Settings in the Miro UI please review [this mapping table](https://docs.google.com/spreadsheets/d/1pd9WuR_7XWVg84h8c7I3kxaYmYnvNKlAcT5f9M5hvUg/edit?usp=sharing). Please make sure you confirm the payload with the official [Miro Teams Settings API documentation](https://developers.miro.com/reference/enterprise-update-team-settings) to make sure you are passing the correct desired values.

5.2. Save your changes.

5.3. In your command line interface navigate to the directory where you have placed the script files (see step 2.2) 

5.4. In your command line interface run `node updateTeams.js`. This command will trigger the script.

## Step 6. Review the folder `miro_teams_update_output_files`

Within the same folder where this script lives, the script creates a folder called `miro_teams_deletion_output_files`. Within this folder you will find the below reports:

  - `Miro_Teams_Update_Results.csv`: Contains the results of the team deletions in a .CSV file
  - `Miro_Teams_Update_Results.json`: Contains the results of the team deletions in a JSON file
  - `Miro_Teams_Overview.csv`: Contains an overview of all teams within the Miro account in a .CSV file
  - `Miro_Teams_Overview.json`: Contains an overview of all teams within the Miro account in a JSON file
  - `Script_Errors.csv`: Contains API errors (if any) in a CSV file. This report is created only if errors occurred, otherwise the file will not be present
  - `Script_Errors.json`: Contains API errors (if any) in a JSON file. This report is created only if errors occurred, otherwise the file will not be present

## Support

If you have any questions or need assistance setting up this application, please reach out to your Miro Customer Success Manager, Onboarding Consultant, Technical Architect, Solutions Engineer or Technical Account Manager.
