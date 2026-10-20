
import { app } from "@azure/functions";

const GRAPH = "https://graph.microsoft.com/v1.0";
const EVENT_ID = "NHS-PROMPTATHON-001";

async function getToken() {
  const {
    AZURE_TENANT_ID,
    AZURE_CLIENT_ID,
    AZURE_CLIENT_SECRET
  } = process.env;

  if (!AZURE_TENANT_ID || !AZURE_CLIENT_ID ||
      !AZURE_CLIENT_SECRET) {
    throw new Error("Missing authentication settings");
  }

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

  return (await response.json()).access_token;
}

async function getAll(url, token) {
  const results = [];

  while (url) {
    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`
      }
    });

    if (!response.ok) {
      throw new Error(`Graph request failed: ${response.status}`);
    }

    const page = await response.json();
    results.push(...(page.value ?? []));
    url = page["@odata.nextLink"] ?? null;
  }

  return results;
}

app.http("status", {
  methods: ["GET"],
  authLevel: "anonymous",

  handler: async (request, context) => {
    try {
      const site = process.env.SHAREPOINT_SITE_ID;
      if (!site) throw new Error("Missing site ID");

      const token = await getToken();

      const lists = await getAll(
        `${GRAPH}/sites/${site}/lists`,
        token
      );

      const list = lists.find(
        item => item.displayName === "Count Me In - CheckIns"
      );

      if (!list) {
        throw new Error("CheckIns list not found");
      }

      const items = await getAll(
        `${GRAPH}/sites/${site}/lists/${list.id}/items?expand=fields`,
        token
      );

      const arrivals = items
        .filter(item => item.fields?.field_1 === EVENT_ID)
        .sort((a, b) =>
          new Date(a.createdDateTime || 0) -
          new Date(b.createdDateTime || 0)
        );

      const latest = arrivals.at(-1);

      // Title is the standard SharePoint title field.
      // Your Power Automate flow previously saved
      // the submitted name into Title.
      const latestName = String(
        latest?.fields?.Title || ""
      ).trim().slice(0, 60);

      return {
        status: 200,
        headers: {
          "Cache-Control": "no-store",
          "Content-Type": "application/json"
        },
        jsonBody: {
          app: "Count Me In!",
          event: "NHS Promptathon",
          status: "ready",
          count: arrivals.length,
          latestGuest: latest
            ? {
                id: latest.id,
                name: latestName || "Guest",
                number: arrivals.length
              }
            : null,
          updatedAt: new Date().toISOString()
        }
      };
    } catch (error) {
      context.error(error.message);

      return {
        status: 500,
        jsonBody: {
          status: "error",
          message: "Unable to retrieve check-ins"
        }
      };
    }
  }
});


