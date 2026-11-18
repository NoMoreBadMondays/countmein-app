
import { app } from "@azure/functions";
import { TableClient } from "@azure/data-tables";

const UUID_PATTERN =
  /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

app.http("status", {
  methods: ["GET"],
  authLevel: "anonymous",
  route: "status",

  handler: async (request, context) => {
    try {
      // Read the event ID using Azure Functions' query API.
      const eventId = request.query.get("event");

      if (!eventId || !UUID_PATTERN.test(eventId)) {
        return {
          status: 400,
          headers: {
            "Cache-Control": "no-store"
          },
          jsonBody: {
            status: "error",
            message: "A valid event ID is required"
          }
        };
      }

      // Connect to Azure Table Storage.
      const connectionString =
        process.env.COUNTMEIN_STORAGE_CONNECTION;

      if (!connectionString) {
        throw new Error(
          "COUNTMEIN_STORAGE_CONNECTION is missing"
        );
      }

      const eventsTable = TableClient.fromConnectionString(
        connectionString,
        "Events"
      );

      const checkinsTable = TableClient.fromConnectionString(
        connectionString,
        "CheckIns"
      );

      // Find the event.
      let event;

      try {
        event = await eventsTable.getEntity(
          "event",
          eventId
        );
      } catch (error) {
        if (error.statusCode === 404) {
          return {
            status: 404,
            headers: {
              "Cache-Control": "no-store"
            },
            jsonBody: {
              status: "error",
              message: "Event not found"
            }
          };
        }

        throw error;
      }

      // Read all check-ins belonging to this event.
      let count = 0;
      let latest = null;

      const checkins = checkinsTable.listEntities({
        queryOptions: {
          filter: `PartitionKey eq '${eventId}'`
        }
      });

      for await (const item of checkins) {
        count++;

        if (
          !latest ||
          String(item.checkedInAt || "") >
            String(latest.checkedInAt || "")
        ) {
          latest = item;
        }
      }

      // Return the data expected by Count Me In!
      return {
        status: 200,
        headers: {
          "Cache-Control": "no-store",
          "Content-Type": "application/json"
        },
        jsonBody: {
          app: "Count Me In!",
          event: event.name,
          status: event.status,
          count: count,
          latestGuest: latest
            ? {
                id: latest.rowKey,
                name: latest.firstName || "Guest",
                number: count
              }
            : null,
          updatedAt: new Date().toISOString()
        }
      };

    } catch (error) {
      context.error(
        "Unable to retrieve event status",
        error
      );

      return {
        status: 500,
        headers: {
          "Cache-Control": "no-store"
        },
        jsonBody: {
          status: "error",
          message: "Unable to retrieve check-ins"
        }
      };
    }
  }
});

