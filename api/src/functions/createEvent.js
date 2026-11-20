
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

function jsonResponse(status, body) {
  return {
    status,
    headers: {
      "Cache-Control": "no-store",
      "Content-Type": "application/json"
    },
    jsonBody: body
  };
}

app.http("createEvent", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "events",

  handler: async (request, context) => {
    try {
      // No Microsoft sign-in required.

      // Read the event details.
      let body;

      try {
        body = await request.json();
      } catch {
        return jsonResponse(400, {
          error: "A valid JSON request is required"
        });
      }

      // Validate the event name.
      const eventName = String(
        body?.name ?? ""
      ).trim();

      if (
        !eventName ||
        eventName.length > 100
      ) {
        return jsonResponse(400, {
          error: "Event name must be between 1 and 100 characters"
        });
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

      // Create a unique event.
      const eventId = randomUUID();

      // Generate a private organiser token.
      // Keep this for future event management features.
      const organiserToken = randomBytes(32)
        .toString("hex");

      await eventsTable.createEntity({
        partitionKey: "event",
        rowKey: eventId,
        name: eventName,
        ownerId: "anonymous",
        organiserTokenHash: hashToken(organiserToken),
        createdAt: new Date().toISOString(),
        status: "active"
      });

      // Generate the event URLs.
      const displayUrl =
        `${WEBSITE_URL}/?event=${encodeURIComponent(eventId)}`;

      const checkinUrl =
        `${WEBSITE_URL}/checkin?event=${encodeURIComponent(eventId)}`;

      return jsonResponse(201, {
        app: "Count Me In!",
        eventId,
        name: eventName,
        displayUrl,
        checkinUrl,
        organiserToken
      });

    } catch (error) {
      context.error("Event creation failed", error);

      return jsonResponse(500, {
        error: "Unable to create event"
      });
    }
  }
});

