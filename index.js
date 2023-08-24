const fs = require('fs');
const readline = require('readline');
const { google } = require('googleapis');
const { Client, GatewayIntentBits } = require('discord.js');
const { DateTime } = require('luxon');

//Discord
const client = new Client({intents: [GatewayIntentBits.Guilds, GatewayIntentBits.GuildMembers]});
const config = require('./config.json') || {};
//config.json should have  { "discordSecret": "asdf", "spreadsheetId": "zxcv", "guildId": "1234" }

// If modifying these scopes, delete token.json.
const SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];
// The file token.json stores the user's access and refresh tokens, and is
// created automatically when the authorization flow completes for the first
// time.
const TOKEN_PATH = 'token.json';

// Load client secrets from a local file.
fs.readFile('credentials.json', async (err, content) => {
    if (err) return console.log('Error loading client secret file:', err);
    // Authorize a client with credentials, then call the Google Sheets API.
    await authorize(JSON.parse(content), storeRoster);
});

/**
 * Create an OAuth2 client with the given credentials, and then execute the
 * given callback function.
 * @param {Object} credentials The authorization client credentials.
 * @param {function} callback The callback to call with the authorized client.
 */
async function authorize(credentials, callback) {
    const { client_secret, client_id, redirect_uris } = credentials.installed;
    const oAuth2Client = new google.auth.OAuth2(
        client_id, client_secret, redirect_uris[0]);

    // Check if we have previously stored a token.
    fs.readFile(TOKEN_PATH, async (err, token) => {
        if (err) return getNewToken(oAuth2Client, callback);
        oAuth2Client.setCredentials(JSON.parse(token));
        await callback(oAuth2Client);
    });
}

/**
 * Get and store new token after prompting for user authorization, and then
 * execute the given callback with the authorized OAuth2 client.
 * @param {google.auth.OAuth2} oAuth2Client The OAuth2 client to get token for.
 * @param {getEventsCallback} callback The callback for the authorized client.
 */
function getNewToken(oAuth2Client, callback) {
    const authUrl = oAuth2Client.generateAuthUrl({
        access_type: 'offline',
        scope: SCOPES,
    });
    console.log('Authorize this app by visiting this url:', authUrl);
    const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout,
    });
    rl.question('Enter the code from that page here: ', (code) => {
        rl.close();
        oAuth2Client.getToken(code, (err, token) => {
            if (err) return console.error('Error while trying to retrieve access token', err);
            oAuth2Client.setCredentials(token);
            // Store the token to disk for later program executions
            fs.writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
                if (err) return console.error(err);
                console.log('Token stored to', TOKEN_PATH);
            });
            callback(oAuth2Client);
        });
    });
}

/**
 * Main method basically
 * @param {google.auth.OAuth2} auth The authenticated Google OAuth client.
 */
async function storeRoster(auth) {
    //Discord setup
    client.on('error', console.error);
    client.on('ready', async () => {
        console.log("INFO: Discord ready.")
        const sheets = google.sheets({ version: 'v4', auth });
        const doit = async () => {
            const guild = client.guilds.cache.get(config.guildId);
            var roster = [];
            var roles = [];

            await guild.roles.fetch();
            await guild.members.fetch();

            //Process roles
            guild.roles.cache.each(role => {
                roles.push(role.name);
            });

            roles = roles.sort();

            //Process members
            guild.members.cache.each(member => {
                let memberRoles = [];
                member.roles.cache.each(memberRole => {
                    memberRoles.push(memberRole.name);
                });
                let finalMemberRoles = [];
                finalMemberRoles.push(member.user.username);
                finalMemberRoles.push(member.nickname);
                roles.forEach((origRole) => {
                    if (memberRoles.includes(origRole)) {
                        finalMemberRoles.push(1);
                    }
                    else {
                        finalMemberRoles.push(0);
                    }
                });
                roster.push(finalMemberRoles);
            });


            //Upload to Sheets

            //Delete rows 3 and greater and columns C and greater
            let request = {
                // The spreadsheet to apply the updates to.
                spreadsheetId: config.spreadsheetId,  

                resource: {
                    "requests": [
                        {
                            "deleteDimension": {
                                "range": {
                                    "sheetId": 0,
                                    "dimension": "ROWS",
                                    "startIndex": 3,
                                }
                            }
                        },
                        {
                            "deleteDimension": {
                                "range": {
                                    "sheetId": 0,
                                    "dimension": "COLUMNS",
                                    "startIndex": 2
                                }
                            }
                        }
                    ]
                },

                auth: auth
            };

            try {
                const response = (await sheets.spreadsheets.batchUpdate(request)).data;
                //console.log(JSON.stringify(response, null, 2));
            } catch (err) {
                console.error(err);
            }




            //Put the role data
            request = {
                // The ID of the spreadsheet to update.
                spreadsheetId: config.spreadsheetId,  

                // The A1 notation of the values to update.
                range: 'C2:2',  

                // How the input data should be interpreted.
                valueInputOption: 'RAW',  

                insertDataOption: 'OVERWRITE', 

                resource: {
                    "majorDimension": "ROWS",
                    "range": "C2:2",
                    "values": [roles]
                },

                auth: auth
            };

            try {
                const response = (await sheets.spreadsheets.values.append(request)).data;
                //console.log(JSON.stringify(response, null, 2));
            } catch (err) {
                console.error(err);
            }

            //Put the member data
            request = {
                // The ID of the spreadsheet to update.
                spreadsheetId: config.spreadsheetId,  

                // The A1 notation of the values to update.
                range: 'A4',  

                // How the input data should be interpreted.
                valueInputOption: 'RAW',  

                insertDataOption: 'OVERWRITE', 

                resource: {
                    "majorDimension": "ROWS",
                    "range": "A4",
                    "values": roster
                },

                auth: auth
            };

            try {
                const response = (await sheets.spreadsheets.values.append(request)).data;
                //console.log(JSON.stringify(response, null, 2));
            } catch (err) {
                console.error(err);
            }

            //Put the date
            request = {
                spreadsheetId: config.spreadsheetId,  
                range: 'B1',  
                valueInputOption: 'USER_ENTERED',  

                resource: {
                    "majorDimension": "ROWS",
                    "range": "B1",
                    "values": [[DateTime.local().setZone('America/New_York').toLocaleString(DateTime.DATETIME_SHORT)]]
                },

                auth: auth
            };

            try {
                const response = (await sheets.spreadsheets.values.update(request)).data;
                //console.log(JSON.stringify(response, null, 2));
            } catch (err) {
                console.error(err);
            }
        };
        doit();
        setInterval(doit, 60 * 60 * 1000);
    });

    client.login(config.discordSecret);
}
