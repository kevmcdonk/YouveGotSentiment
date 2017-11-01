Before deploying, you must set the following values in src/components/app.tsx:

const textAnalyticsKey = '';
const contentModeratorKey = '';

and also gulpfile.js

var accountName = ''; // Azure Blob Storage account name
var accountKey = ''; // Azure Blob Storage account key

You will need to create the CDN entry, text analytics and content moderator in the Azure Portal. You can read more at:

