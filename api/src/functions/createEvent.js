
import { app } from "@azure/functions";
import { TableClient } from "@azure/data-tables";
import { randomBytes, randomUUID, createHash } from "node:crypto";

const WEBSITE_URL =
  "https://ambitious-bay-07ef57e10.1.azurestaticapps.net";

function hashToken(token) {
  return createHash("sha256")
    .update(token)
    .digest("hex");
}

app.http("createEvent", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "events",

  handler: async (request, context) => {
    try {
      // Confirm the organiser is signed in.
      const encoded = request.headers.get(
        "x-ms-client-principal"
      );

      if (!encoded) {
        return {
          status: 401,
          jsonBody: {
            error: "Organiser sign-in required"
          }
        };
      }

      let principal;

      try {
        principal = JSON.parse(
          Buffer.from(encoded, "base64").toString("utf8")
        );
      } catch {
        return {
          status: 401,
          jsonBody: {
            error: "Invalid organiser identity"
          }
        };
      }

      if (
        !principal.userId ||
        !principal.identityProvider ||
        !principal.userRoles?.includes("authenticated")
      ) {
        return {
          status: 401,
          jsonBody: {
            error: "Authentication required"
          }
        };
      }

      // Read and validate the event name.
      let body;

      try {
        body = await request.json();
      } catch {
        return {
          status: 400,
          jsonBody: {
            error: "A valid JSON request is required"
          }
        };
      }

      const eventName = String(
        body?.name ?? ""
      ).trim();

      if (
        !eventName ||
        eventName.length > 100
      ) {
        return {
          status: 400,
          jsonBody: {
            error: "Event name must be between 1 and 100 characters"
          }
        };
      }

      // Connect to Azure Table Storage.
      const connectionString =
        process.env.COUNTMEIN_STORAGE_CONNECTION;

      if (!connectionString) {
        throw new Error(
          "Missing Azure Storage connection setting"
        );
      }

      const eventsTable = TableClient.fromConnectionString(
        connectionString,
        "Events"
      );

      // Create a unique event and private organiser token.
      const eventId = randomUUID();

      const organiserToken = randomBytes(32)
        .toString("hex");

      const ownerId =
        `${principal.identityProvider}:${principal.userId}`;

      await eventsTable.createEntity({
        partitionKey: "event",
        rowKey: eventId,
        name: eventName,
        ownerId,
        organiserTokenHash: hashToken(organiserToken),
        createdAt: new Date().toISOString(),
        status: "active"
      });

      // Generate both public-facing event URLs.
      const displayUrl =
        `${WEBSITE_URL}/?event=${encodeURIComponent(eventId)}`;

      const checkinUrl =
        `${WEBSITE_URL}/checkin?event=${encodeURIComponent(eventId)}`;

      return {
        status: 201,
        headers: {
          "Cache-Control": "no-store",
          "Content-Type": "application/json"
        },
        jsonBody: {
          app: "Count Me In!",
          eventId,
          name: eventName,
          displayUrl,
          checkinUrl,
          organiserToken
        }
      };

    } catch (error) {
      context.error("Event creation failed", error);

      return {
        status: 500,
        jsonBody: {
          error: "Unable to create event"
        }
      };
    }
  }
});

