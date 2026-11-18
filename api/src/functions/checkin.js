
import { app } from "@azure/functions";
import { TableClient } from "@azure/data-tables";
import { randomUUID } from "node:crypto";

app.http("checkin", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "checkin",

  handler: async (request, context) => {
    try {
      const connectionString =
        process.env.COUNTMEIN_STORAGE_CONNECTION;

      if (!connectionString) {
        throw new Error("Storage connection is missing");
      }

      let body;

      try {
        body = await request.json();
      } catch {
        return {
          status: 400,
          jsonBody: { error: "Invalid request" }
        };
      }

      const eventId = String(body?.eventId ?? "").trim();
      const firstName = String(body?.firstName ?? "").trim();

      if (
        !/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(eventId) ||
        !firstName ||
        firstName.length > 60 ||
        /[\x00-\x1f\x7f]/.test(firstName)
      ) {
        return {
          status: 400,
          jsonBody: {
            error: "Invalid event or first name"
          }
        };
      }

      const eventsTable = TableClient.fromConnectionString(
        connectionString,
        "Events"
      );

      const checkinsTable = TableClient.fromConnectionString(
        connectionString,
        "CheckIns"
      );

      // Confirm the event exists and is active.
      let event;

      try {
        event = await eventsTable.getEntity("event", eventId);
      } catch (error) {
        if (error.statusCode === 404) {
          return {
            status: 404,
            jsonBody: { error: "Event not found" }
          };
        }
        throw error;
      }

      if (event.status !== "active") {
        return {
          status: 403,
          jsonBody: { error: "This event is closed" }
        };
      }

      const checkinId = randomUUID();

      await checkinsTable.createEntity({
        partitionKey: eventId,
        rowKey: checkinId,
        firstName,
        checkedInAt: new Date().toISOString()
      });

      return {
        status: 201,
        headers: {
          "Cache-Control": "no-store"
        },
        jsonBody: {
          success: true,
          message: "You're checked in!"
        }
      };

    } catch (error) {
      context.error("Check-in failed", error);

      return {
        status: 500,
        jsonBody: {
          error: "Unable to complete check-in"
        }
      };
    }
  }
});
