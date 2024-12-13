/// <summary>
/// This class is responsible for handling the messaging extension code and SSO auth inside copilot.
/// </summary>

const { TeamsActivityHandler, CardFactory } = require("botbuilder");
const config = require("./config");
const azure = require("azure-storage");

const { Client } = require("@microsoft/microsoft-graph-client");
const { TokenCredentialAuthenticationProvider } = require("@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials");
const {
  MessageExtensionTokenResponse,
  handleMessageExtensionQueryWithSSO,
  OnBehalfOfCredentialAuthConfig,
  OnBehalfOfUserCredential
} = require("@microsoft/teamsfx");

const { OnBehalfOfCredential, DeviceCodeCredential } = require("@azure/identity");


class SearchApp extends TeamsActivityHandler {

  async handleTeamsMessagingExtensionQuery(context, query) {
    const { parameters } = query;

    // const file = getParameterByName(parameters, "File");
    const skills = getParameterByName(parameters, "Skill");
    const country = getParameterByName(parameters, "Location");
    const availabilityParam = getParameterByName(parameters, "Availability");
    
    var availability;

    if (availabilityParam == "true") {
      availability = true;
    }
    else if (availabilityParam == "false") {
      availability = false;
    }
    else {
      availability = undefined;
    }

    function constructSearchObject(skills, country, availability) {
      const filterObject = {};

      if (country) {
        filterObject.country = country;
      }

      if (skills) {
        filterObject.skills = skills;
      }

      if (availability != undefined) {
        filterObject.availability = availability;
      }

      return filterObject;
    }

    const searchObject = constructSearchObject(skills, country, availability);

    // Define your Azure Table Storage connection string or credentials
    const storageConnectionString = config.storageConnectionString;

    // Create a table service object using the connection string
    const tableService = azure.createTableService(storageConnectionString);

    var candidateData = [];

    // Define the name of the table you want to store data in
    const tableName = config.storageTableName;
    
    // When the Bot Service Auth flow completes, the query.State will contain a magic code used for verification.
    const magicCode =
      query.state && Number.isInteger(Number(query.state))
        ? query.state
        : '';

    const tokenResponse = await context.adapter.getUserToken(
      context,
      "authbot",
      magicCode
    );

    if (!tokenResponse || !tokenResponse.token) {
      // There is no token, so the user has not signed in yet.
      // Retrieve the OAuth Sign in Link to use in the MessagingExtensionResult Suggested Actions
      const signInLink = await context.adapter.getSignInLink(
        context,
        "authbot"
      );
      consloe.log("signInLink:", signInLink);

      return {
        composeExtension: {
          type: 'auth',
          suggestedActions: {
            actions: [
              {
                type: 'openUrl',
                title: 'Bot Service OAuth',
                value: signInLink
              },
            ],
          },
        },
      }      
    };

    // Initialize DeviceCodeCredential
    const credential = new DeviceCodeCredential({
      tenantId: config.tenantId,
      clientId: config.clientId,
      userPromptCallback: (info) => {
        console.log(info.message); // Display the device code message to the user
      },
    });

    // Create authentication provider for Microsoft Graph
    const authProvider = new TokenCredentialAuthenticationProvider(credential, {
      scopes: ['User.Read'], // Specify the required scopes
    });

    // Initialize Microsoft Graph client
    const graphClient = Client.initWithMiddleware({ authProvider: authProvider });
    
    async function fetchCandidatesFromGraph(graphClient, filters) {
      const queryParams = [];
    
      if (filters.skills) {
        queryParams.push(`skills eq '${filters.skills}'`); // Replace with your attribute
      }
      if (filters.country) {
        queryParams.push(`country eq '${filters.country}'`); // Replace with your attribute
      }
      if (filters.availability !== undefined) {
        queryParams.push(`availability eq ${filters.availability}`); // Replace with your attribute
      }
    
      const filterQuery = queryParams.length > 0 ? queryParams.join(" and ") : "";
    
      try {
        const users = await graphClient
          .api(`/users`)
          .filter(filterQuery)
          .select("id,displayName,skills,country,availability") // Specify required fields
          .get();
    
        return users.value;
      } catch (error) {
        console.error("Error fetching candidates from Microsoft Graph:", error);
        return [];
      }
    }

    const candidatesFromGraph = await fetchCandidatesFromGraph(graphClient, searchObject);



    // Define a function to fetch candidates based on parameters
    function fetchCandidates(queryParameters) {
      return new Promise((resolve, reject) => {
        const query = new azure.TableQuery();

        let whereClause = "";
        let skillsAdded = false;

        // Construct the where clause dynamically based on provided parameters
        Object.keys(queryParameters).forEach((key, index) => {
          if (key === "skills" || key === "availability") {
            return; // Skip skills and availability for now, handle separately below
          }

          const condition = `${key} eq '${queryParameters[key]}'`;
          if (whereClause !== "") {
            whereClause += " and ";
          }
          whereClause += `(${condition})`;
        });

        // Add availability filter if provided
        if (queryParameters.availability !== undefined && queryParameters.availability !== null) {
          const availabilityCondition = `availability eq ${queryParameters.availability}`;
          if (whereClause !== "") {
            whereClause += " and ";
          }
          whereClause += `(${availabilityCondition})`;
        }

        // If no parameters provided, select all
        if (whereClause === "") {
          whereClause = "PartitionKey ne ''"; // Dummy condition to select all in case parameters are null or empty
        }

        query.where(whereClause);

        tableService.queryEntities(
          tableName,
          query,
          null,
          (error, result, response) => {
            if (error) {
              reject(error);
              return;
            }

            let filteredCandidates = result.entries;

            // Filter candidates based on skills
            if (queryParameters.skills) {
              const skills = queryParameters.skills
                .split(",")
                .map((skill) => skill.trim().toLowerCase());
              filteredCandidates = filteredCandidates.filter((candidate) => {
                const candidateSkills = candidate.skills._.split(",").map(
                  (skill) => skill.trim().toLowerCase()
                );
                return candidateSkills.some((candidateSkills) => candidateSkills.includes(skills));
              });
            }

            resolve(filteredCandidates);
          }
        );
      });
    }

    // Fetch candidates based on applied filters.
    var candidates = await fetchCandidates(searchObject);

    var attachments = [];
    // candidateData = candidates;
    candidateData = candidatesFromGraph
    console.log("Candidates:", candidateData);

    // Create Adaptive Card object
    candidateData.map((result) => {
      var availability = result.availability._ ? "Yes" : "No"
      const resultCard = CardFactory.adaptiveCard({
        "type": "AdaptiveCard",
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "version": "1.4",
        "body": [
          {
            "type": "TextBlock",
            "text": "Expert Finder",
            "wrap": true,
            "size": "Large",
            "weight": "Bolder",
            "separator": true
          },
          {
            "type": "ColumnSet",
            "columns": [
              {
                "type": "Column",
                "items": [
                  {
                    "type": "Image",
                    "url": "https://pbs.twimg.com/profile_images/3647943215/d7f12830b3c17a5a9e4afcc370e3a37e_400x400.jpeg",
                    "altText": "profileImage",
                    "size": "Small",
                    "style": "Person"
                  }
                ],
                "width": "auto"
              },
              {
                "type": "Column",
                "items": [
                  {
                    "type": "TextBlock",
                    "weight": "Bolder",
                    "text": `${result.name._}`,
                    "wrap": true,
                    "spacing": "None",
                    "horizontalAlignment": "Left",
                    "maxLines": 0,
                    "size": "Medium"
                  }
                ],
                "width": "stretch",
                "spacing": "Medium",
                "verticalContentAlignment": "Center"
              }
            ]
          },
          {
            "type": "FactSet",
            "facts": [
              {
                "title": "Skills:",
                "value": `${result.skills._}`
              },
              {
                "title": "Location:",
                "value": `${result.country._}`,
              },
              {
                "title": "Available:",
                "value": `${availability}`,
              }
            ]
          }
        ],
        // New actions
        "actions": [
          {
              "type": "Action.OpenUrl",
              "title": "Action Open URL",
              "url": "https://adaptivecards.io"
          },
          {
              "type": "Action.ShowCard",
              "title": "Action Submit",
              "card": {
                  "type": "AdaptiveCard",
                  "version": "1.5",
                  "body": [
                      {
                          "type": "Input.Text",
                          "id": "name",
                          "label": "Please enter your name:",
                          "isRequired": true,
                          "errorMessage": "Name is required"
                      }
                  ],
                  "actions": [
                      {
                          "type": "Action.Submit",
                          "title": "Submit"
                      }
                  ]
              }
          },
          {
              "type": "Action.ShowCard",
              "title": "Action ShowCard",
              "card": {
                  "type": "AdaptiveCard",
                  "version": "1.0",
                  "body": [
                      {
                          "type": "TextBlock",
                          "text": "This card's action will show another card"
                      }
                  ],
                  "actions": [
                      {
                          "type": "Action.ShowCard",
                          "title": "Action.ShowCard",
                          "card": {
                              "type": "AdaptiveCard",
                              "body": [
                                  {
                                      "type": "TextBlock",
                                      "text": "**Welcome To New Card**"
                                  },
                                  {
                                      "type": "TextBlock",
                                      "text": "This is your new card inside another card"
                                  }
                              ]
                          }
                      }
                  ]
              }
          }
      ]
        
      });

      const previewCard = CardFactory.heroCard(
        result.name._,
        result.skills._
      );

      attachments.push({ ...resultCard, preview: previewCard });
    });

    return {
      composeExtension: {
        type: "result",
        attachmentLayout: "list",
        attachments: attachments,
      },
    };
  }
}

const getParameterByName = (parameters, name) => {
  const param = parameters.find((p) => p.name === name);
  return param ? param.value : "";
};

module.exports = { SearchApp };