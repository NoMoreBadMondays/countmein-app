
import { app } from '@azure/functions';

app.http('status', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'status',

  handler: async () => {
    return {
      status: 200,
      jsonBody: {
        app: 'Count Me In!',
        event: 'NHS Promptathon',
        status: 'ready',
        message: 'Azure API is working!'
      }
    };
  }
});
