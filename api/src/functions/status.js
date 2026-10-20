
import { app } from "@azure/functions";

const GRAPH = "https://graph.microsoft.com/v1.0";
const EVENT_ID = "NHS-PROMPTATHON-001";

async function getAccessToken() {
  const {
    AZURE_TENANT_ID,
    AZURE_CLIENT_ID,
    AZURE_CLIENT_SECRET
  } = process.env;

  const response = await fetch(
    `https://login.microsoftonline.com/${AZURE_TENANT_ID}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded"
      },
      body: new URLSearchParams({
        client_id: AZURE_CLIENT_ID,
        client_secret: AZURE_CLIENT_SECRET,
        scope: "https://graph.microsoft.com/.default",
        grant_type: "client_credentials"
      })
    }
  );

  if (!response.ok) {
    throw new Error(`Authentication failed: ${response.status}`);
  }

  const data = await response.json();
  return data.access_token;
}

async function graphGet(path, token) {
  const response = await fetch(`${GRAPH}${path}`, {
    headers: { Authorization: `Bearer ${token}` }
  });

  if (!response.ok) {
    throw new Error(`Microsoft Graph returned ${response.status}`);
  }

  return response.json();
}

async function graphGetAll(path, token) {
  const items = [];
  let next = `${GRAPH}${path}`;

  while (next) {
    const response = await fetch(next, {
      headers: { Authorization: `Bearer ${token}` }
    });

    if (!response.ok) {
      throw new Error(`Microsoft Graph returned ${response.status}`);
    }

    const page = await response.json();
    items.push(...(page.value ?? []));
    next = page["@odata.nextLink"] ?? null;
  }

  return items;
}

app.http("status", {
  methods: ["GET"],
  authLevel: "anonymous",
  handler: async (request, context) => {
    try {
      const siteId = process.env.SHAREPOINT_SITE_ID;
      const token = await getAccessToken();

      // Find the CheckIns list by its display name.
      const lists = await graphGetAll(
        `/sites/${siteId}/lists`,
        token
      );

      const checkInsList = lists.find(
        (list) => list.displayName === "Count Me In - CheckIns"
      );

      if (!checkInsList) {
        throw new Error("CheckIns list not found");
      }

      // Read all pages, including SharePoint column values.
      const items = await graphGetAll(
        `/sites/${siteId}/lists/${checkInsList.id}/items?expand=fields`,
        token
      );

      const arrivals = items.filter(
        (item) => item.fields?.EventID === EVENT_ID
      );

      return {
        status: 200,
        headers: {
          "Content-Type": "application/json",
          "Cache-Control": "no-store"
        },
        jsonBody: {
          app: "Count Me In!",
          event: "NHS Promptathon",
          status: "ready",
          count: arrivals.length,
          updatedAt: new Date().toISOString()
        }
      };
    } catch (error) {
      context.error("Count Me In status error:", error.message);

      return {
        status: 500,
        jsonBody: {
          status: "error",
          message: "Unable to retrieve live check-ins"
        }
      };
    }
  }
});
