#r "Newtonsoft.Json"
using System;
using System.Net;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
 
public static async Task<HttpResponseMessage> Run(HttpRequestMessage req, TraceWriter log)
{
    bool PRODUCTION = true;

    string TENANT_ID_FIELD, CUSTOMER_EMAIL_FIELD, BACKUP_EMAIL_FIELD;
    Dictionary<string, string> teams_hook_uris, domain_tenants;

    if (PRODUCTION) {
        TENANT_ID_FIELD = "MicrosoftTeamsCMMI.CustomerName";
        CUSTOMER_EMAIL_FIELD = "MicrosoftTeamsCMMI.CustomerEmail";
        BACKUP_EMAIL_FIELD = "Microsoft.VSTS.TCM.ReproSteps";

        teams_hook_uris = new Dictionary<string, string>()
        {
            // A dictionary where teams_hook_uris[tenantId] = webhook_uri.
            { "aaaa-aaaa-aaaa", "https://outlook.office.com/webhook/15f69bfd-b260-478c-af37-db567141623a@240a3177-7696-41df-a4ea-0d1b0999fb38/IncomingWebhook/ced5606707b04b5c933fc0b09979889e/2c46ba10-5141-459c-a866-d47436de0cba"},
            { "bbbb-bbbb-bbbb", "https://outlook.office.com/webhook/15f69bfd-b260-478c-af37-db567141623a@240a3177-7696-41df-a4ea-0d1b0999fb38/IncomingWebhook/ca9d8272d45a466592caff1e7668ff4b/2c46ba10-5141-459c-a866-d47436de0cba"},
            { "cccc-cccc-cccc", "https://outlook.office.com/webhook/15f69bfd-b260-478c-af37-db567141623a@240a3177-7696-41df-a4ea-0d1b0999fb38/IncomingWebhook/c487e2afce904e5a833bd7fa114185b5/2c46ba10-5141-459c-a866-d47436de0cba"},
            { "dddd-dddd-dddd", "https://outlook.office.com/webhook/15f69bfd-b260-478c-af37-db567141623a@240a3177-7696-41df-a4ea-0d1b0999fb38/IncomingWebhook/4d7cdaf9cd5a42fbb77cb8d43050b7ee/2c46ba10-5141-459c-a866-d47436de0cba"},
            { "eeee-eeee-eeee", "https://outlook.office.com/webhook/15f69bfd-b260-478c-af37-db567141623a@240a3177-7696-41df-a4ea-0d1b0999fb38/IncomingWebhook/03e6db947cf14ae5965deb2b2f25145d/2c46ba10-5141-459c-a866-d47436de0cba"},
        };
    } else {
        TENANT_ID_FIELD = "Custom.CustomerName";
        CUSTOMER_EMAIL_FIELD = "Custom.CustomerEmail";
        teams_hook_uris = new Dictionary<string, string>() {
            { "aaaa-aaaa-aaaa", "https://outlook.office.com/webhook/15f69bfd-b260-478c-af37-db567141623a@240a3177-7696-41df-a4ea-0d1b0999fb38/IncomingWebhook/ced5606707b04b5c933fc0b09979889e/2c46ba10-5141-459c-a866-d47436de0cba"},
            { "bbbb-bbbb-bbbb", "https://outlook.office.com/webhook/15f69bfd-b260-478c-af37-db567141623a@240a3177-7696-41df-a4ea-0d1b0999fb38/IncomingWebhook/ca9d8272d45a466592caff1e7668ff4b/2c46ba10-5141-459c-a866-d47436de0cba"},
            { "cccc-cccc-cccc", "https://outlook.office.com/webhook/15f69bfd-b260-478c-af37-db567141623a@240a3177-7696-41df-a4ea-0d1b0999fb38/IncomingWebhook/c487e2afce904e5a833bd7fa114185b5/2c46ba10-5141-459c-a866-d47436de0cba"},
            { "dddd-dddd-dddd", "https://outlook.office.com/webhook/15f69bfd-b260-478c-af37-db567141623a@240a3177-7696-41df-a4ea-0d1b0999fb38/IncomingWebhook/4d7cdaf9cd5a42fbb77cb8d43050b7ee/2c46ba10-5141-459c-a866-d47436de0cba"},
            { "eeee-eeee-eeee", "https://outlook.office.com/webhook/15f69bfd-b260-478c-af37-db567141623a@240a3177-7696-41df-a4ea-0d1b0999fb38/IncomingWebhook/03e6db947cf14ae5965deb2b2f25145d/2c46ba10-5141-459c-a866-d47436de0cba"},
        };
    }

    string TEAMS_HOOK_URI;

    log.Info("Webhook was triggered!");
    HttpContent requestContent = req.Content;
    string jsonContent = requestContent.ReadAsStringAsync().Result;

    log.Info(jsonContent);

    dynamic jsonObj = JObject.Parse(jsonContent);

    // Quickly figure out if the Tenant ID is in there
    string customerId = "";
    string eventType = jsonObj.eventType;
    if (jsonObj.eventType == "workitem.updated") {
        customerId = jsonObj.resource.revision.fields[TENANT_ID_FIELD];
    } else if (jsonObj.eventType == "workitem.created") {
        customerId = jsonObj.resource.fields[TENANT_ID_FIELD];
    }

    try {
        TEAMS_HOOK_URI = teams_hook_uris[customerId];
    } catch (Exception e) {
        // TODO: Try/catch parsing the repro steps here
        try {
            string reproSteps = jsonObj.resource.revision.fields["Microsoft.VSTS.TCM.ReproSteps"];
            string userEmail = reproSteps.Substring(reproSteps.IndexOf("<b>User's email address</b>:") + 28);
            log.Info("userEmail is: " + userEmail);

            // TODO: Check if it ends with one of the domains, try to assign it a tenantID, then exit this
        } catch (Exception e2) {
            log.Info("Exception was: " + e2);
            log.Info("Nothing in the repro steps either");
        }

        log.Info("This tenant wasn't found: " + customerId + " Not notifying anyone");
        return req.CreateResponse(HttpStatusCode.OK);
    }
    log.Info("The tenant is " + customerId);
    
    int bugId;
    string bugState, bugReason, bugTitle, bugSubmitter, bugSteps, bugSeverity, bugLink, bugLinkString, stateString;

    stateString = "";
    customerId = "";
    bugTitle = "";
    bugSubmitter = "";
    bugReason = "";
    bugSteps = "";
    bugLink = "";


    if (eventType == "workitem.updated") {
        bugState = jsonObj.resource.revision.fields["System.State"];
        bugId = jsonObj.resource.revision.id;
        if ((bugState != "Resolved") && (bugState != "Closed")) {
            log.Info("Bug state was " + bugState + ", not notifying anyone");
            return req.CreateResponse(HttpStatusCode.OK);
        }

        bugTitle = jsonObj.resource.revision.fields["System.Title"];
        bugLink = jsonObj.resource["_links"]["html"]["href"];
        bugLinkString = $"**Link to VSTS Item**: [Bug {bugId}]({bugLink})";
        bugSubmitter = jsonObj.resource.revision.fields[CUSTOMER_EMAIL_FIELD];
        bugReason = jsonObj.resource.revision.fields["System.Reason"];
        bugSteps = bugLinkString + "\n\n" + jsonObj.resource.revision.fields["Microsoft.VSTS.TCM.ReproSteps"];
        
        

        stateString = "Bug " + bugId.ToString() + " was " + bugState;
    } 
    else if (eventType == "workitem.created") {
        bugState = jsonObj.resource.fields["System.State"];  // always "New"
        bugTitle = jsonObj.resource.fields["System.Title"];
        bugId = jsonObj.resource.id;
        bugLink = jsonObj.resource["_links"]["html"]["href"];
        bugLinkString = $"**Link to VSTS Item**: [Bug {bugId}]({bugLink})";
        
        bugSubmitter = jsonObj.resource.fields[CUSTOMER_EMAIL_FIELD];
        bugSteps = bugLinkString + "\n\n" + jsonObj.resource.fields["Microsoft.VSTS.TCM.ReproSteps"];
        

        stateString = "Bug " + bugId.ToString() + " was created";
    } else {
        log.Info("Whoops, that wasn't a create or update event. This shouldn't happen!");
        return req.CreateResponse(HttpStatusCode.OK);
    }

    // Generate the card that will show up in Teams
    log.Info("About to make betterCard");

    dynamic card = new {
        summary = stateString,
        title = stateString,
        sections = new object [] {
            new {
                activityTitle = bugTitle,
                activitySubtitle = bugSteps,
                activityImage = "https://i.imgur.com/xqG1HMv.png",
                facts = new object [] {
                    new {
                        name = "Submitted By",
                        value = bugSubmitter,
                    },
                    new {
                        name = "State",
                        value = bugState
                    },
                    new {
                        name = "Reason",
                        value = bugReason
                    }
                }
            }
        }
    };

    JObject cardJObject = JObject.FromObject(card);
    string cardJson = cardJObject.ToString();
    log.Info(cardJson);

    using (var client = new HttpClient())
    {
        var response = await client.PostAsync(
            TEAMS_HOOK_URI,
            new StringContent(cardJson, System.Text.Encoding.UTF8, "application/json")
        );
    }

    return req.CreateResponse(HttpStatusCode.OK);
}
