
import { app } from "@azure/functions";
import { TableClient } from "@azure/data-tables";
import { randomBytes, randomUUID, createHash } from "node:crypto";

const storage = process.env.COUNTMEIN_STORAGE_CONNECTION;

function makeEventId() {
  return randomUUID();
}

function hashToken(token) {
  return createHash("sha256").update(token).digest("hex");
}

app.http("createEvent", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "events",

  handler: async (request, context) => {
    try {
      // Only authenticated organisers may create events.
      // Azure Static Web Apps provides identity information.
      const encoded = request.headers.get("x-ms-client-principal");

      if (!encoded) {
        return {
          status: 401,
          jsonBody: { error: "Sign-in required" }
        };
      }

      const principal = JSON.parse(
        Buffer.from(encoded, "base64").toString("utf8")
      );

      if (!principal.userId || !principal.identityProvider) {
        return {
          status: 401,
          jsonBody: { error: "Invalid identity" }
        };
      }

      const body = await request.json();
      const eventName = String(body?.name ?? "").trim();

      if (!eventName || eventName.length > 100) {
        return {
          status: 400,
          jsonBody: {
            error: "Event name must be between 1 and 100 characters"
          }
        };
      }

      if (!storage) {
        throw new Error("Storage connection is not configured");
      }

      const eventId = makeEventId();
      const organiserToken = randomBytes(32).toString("hex");

      const client = TableClient.fromConnectionString(
        storage,
        "Events"
      );

      await client.createEntity({
        partitionKey: "event",
        rowKey: eventId,
        name: eventName,
        ownerId: `${principal.identityProvider}:${principal.userId}`,
        organiserTokenHash: hashToken(organiserToken),
        createdAt: new Date().toISOString(),
        status: "active"
      });

      const baseUrl = new URL(request.url).origin;

      return {
        status: 201,
        headers: {
          "Cache-Control": "no-store"
        },
        jsonBody: {
          eventId,
          name: eventName,
          displayUrl: `${baseUrl}/?event=${eventId}`,
          checkinUrl: `${baseUrl}/checkin?event=${eventId}`,
          organiserToken
        }
      };
    } catch (error) {
      context.error("Event creation failed", error);

      return {
        status: 500,
        jsonBody: { error: "Unable to create event" }
      };
    }
  }
});
