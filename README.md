# Site Control OOH - React + Node.js + MySQL Web App

This is the standalone web-app conversion of the Site Control OOH WordPress plugin. WordPress/PHP is no longer required at runtime.

## Stack
- React 18 + Vite frontend
- Node.js + Express REST API
- Hostinger MySQL / MariaDB using `mysql2`
- JWT login and role-ready user table
- Multer image uploads
- SheetJS Excel import/export
- PptxGenJS automated PowerPoint generation

## Preserved operational modules
Dashboard, Sites, Campaign Tracker, Occupancy, Proposal Builder, Electricity, Vendors, Clients, Invoices, Import/Export, Reports, Notifications, Activity Log, Settings, site image uploads and Automated PPT.

The Automated PPT retains the approved Media Buzz sequence and layout: fixed page 1 -> fixed page 2 -> selected site image slides -> fixed page 3. Each selected site can have multiple images, producing one full right-side image slide per image. Site details remain in the left yellow/navy panel with bullet-aligned rows, optional Rate, and separate Latitude/Longitude rows.

## Hostinger deployment
1. In hPanel create a MySQL database and user. Open phpMyAdmin and import `sql/schema.sql`.
2. Edit root `.env` with the exact Hostinger DB host, database, username and password. Set `APP_URL` and `CLIENT_URL` to your real HTTPS domain. Change `JWT_SECRET` and initial admin credentials.
3. Install Node packages:
   `npm run install:all`
4. Build React:
   `npm run build`
5. Start Node API:
   `npm start`
6. In Hostinger Node.js App settings use project root as application root and `server/src/index.js` as startup file. Ensure the app has write permission to `server/uploads`.
7. If your hosting uses a different external port, keep the platform-provided `PORT`; the app reads `process.env.PORT`.

## First login
On the first server start, if the `users` table is empty, the app creates the administrator from `ADMIN_EMAIL` and `ADMIN_PASSWORD` in `.env`.

## Existing plugin data migration
Export the plugin's Sites Excel and import it from the web app's Import / Export screen. `PPT Availability`, `Selling Amount`, Latitude and Longitude are mapped. For a full historical migration of all WordPress plugin tables, export the WordPress `wp_scooh_*` tables and rename them to the corresponding standalone table names in this schema before import.

## Development
- API: `cd server && npm install && npm run dev`
- UI: `cd client && npm install && npm run dev`
- Vite development requests should proxy `/api` to port 3000 if developing separately; production is served by Express from `client/dist`.
