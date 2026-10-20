
import { app } from "@azure/functions";

const GRAPH = "https://graph.microsoft.com/v1.0";
const EVENT_ID = "NHS-PROMPTATHON-001";

async function getToken() {
  const tenant = process.env.AZURE_TENANT_ID;
  const client = process.env.AZURE_CLIENT_ID;
  const secret = process.env.AZURE_CLIENT_SECRET;

  if (!tenant || !client || !secret) {
    throw new Error("Missing Azure authentication settings");
  }

  const response = await fetch(
    `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded"
      },
      body: new URLSearchParams({
        client_id: client,
        client_secret: secret,
        grant_type: "client_credentials",
        scope: "https://graph.microsoft.com/.default"
      })
    }
  );

  if (!response.ok) {
    throw new Error(`Authentication failed (${response.status})`);
  }

  return (await response.json()).access_token;
}

async function getAll(url, token) {
  const results = [];

  while (url) {
    const response = await fetch(url, {
      headers: { Authorization: `Bearer ${token}` }
    });

    if (!response.ok) {
      throw new Error(`Graph request failed (${response.status})`);
    }

    const data = await response.json();
    results.push(...(data.value ?? []));
    url = data["@odata.nextLink"];
  }

  return results;
}

app.http("status", {
  methods: ["GET"],
  authLevel: "anonymous",
  handler: async (request, context) => {
    try {
      const site = process.env.SHAREPOINT_SITE_ID;

      if (!site) {
        throw new Error("Missing SHAREPOINT_SITE_ID");
      }

      const token = await getToken();

      const lists = await getAll(
        `${GRAPH}/sites/${site}/lists`,
        token
      );

      const checkIns = lists.find(
        list => list.displayName === "Count Me In - CheckIns"
      );

      if (!checkIns) {
        throw new Error("CheckIns list not found");
      }

      const items = await getAll(
        `${GRAPH}/sites/${site}/lists/${checkIns.id}/items?expand=fields`,
        token
      );

      // field_1 is the internal SharePoint name for EventID.
      const count = items.filter(
        item => item.fields?.field_1 === EVENT_ID
      ).length;

      return {
        status: 200,
        headers: {
          "Cache-Control": "no-store"
        },
        jsonBody: {
          app: "Count Me In!",
          event: "NHS Promptathon",
          status: "ready",
          count,
          updatedAt: new Date().toISOString()
        }
      };
    } catch (error) {
      context.error(error.message);

      return {
        status: 500,
        jsonBody: {
          status: "error",
          message: "Could not retrieve check-ins"
        }
      };
    }
  }
});

