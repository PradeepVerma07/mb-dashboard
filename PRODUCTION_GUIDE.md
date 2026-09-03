# Media Buzz OOH Workspace — Production Guide & System Documentation
**Document Version:** 3.0.0  
**Target Environment:** Production (Hostinger / VPS / AWS / Cloud)  
**Applicable Platform:** Node.js (v18+) + Express + React (v18) + MySQL 8.0  

---

## Table of Contents
1. [System Architecture & Overview](#1-system-architecture--overview)
2. [Complete Functional Module Manual](#2-complete-functional-module-manual)
3. [Production Advantages & Strategic Value](#3-production-advantages--strategic-value)
4. [Production Disadvantages, Trade-offs & Mitigations](#4-production-disadvantages-trade-offs--mitigations)
5. [Production Deployment & Hostinger Setup Checklist](#5-production-deployment--hostinger-setup-checklist)
6. [Security, Backups & Disaster Recovery](#6-security-backups--disaster-recovery)
7. [API & Data Reference](#7-api--data-reference)

---

## 1. System Architecture & Overview

### 1.1 Architecture Blueprint
```
+-----------------------------------------------------------------------+
|             Client Browser (React 18 Single Page App / Vite)          |
|  - Real-time Operations Dashboard    - Automated PPT Generator        |
|  - Proposal Letterhead Engine        - Exact Excel Import/Export      |
|  - User Management & RBAC            - Meter & Campaign Trackers      |
+-----------------------------------+-----------------------------------+
                                    | HTTPS / REST API
                                    v
+-----------------------------------------------------------------------+
|                    Express Backend API Server (Node.js)               |
|  - JWT Stateless Auth & Bcrypt    - Multer Multi-Image Uploads        |
|  - Dynamic Route Hierarchy        - Excel & JSON Stream Parsers       |
+-------------------+-----------------------------------+---------------+
                    |                                   |
                    v                                   v
+-----------------------------+       +---------------------------------+
|   MySQL 8.0 Database        |       | Public Upload Storage (/uploads)|
|   (InnoDB Structured Schema)|       | (High-Res Photos & Slides)      |
+-----------------------------+       +---------------------------------+
```

### 1.2 Technology Stack
- **Frontend Layer:** React 18, React Router v6, Vite build tool, modern CSS grid/flexbox dark design system.
- **Backend API Layer:** Express.js REST API with JWT (JSON Web Tokens) stateless authentication, `bcrypt` password hashing, `multer` multipart upload handling, and parameterized SQL queries.
- **Database Layer:** MySQL 8.0 (InnoDB) with foreign key constraints, composite indexing, and `DATETIME` audit timestamps.
- **Static Asset Serving:** Express static file streaming for high-resolution hoarding photographs and cover slides.

---

## 2. Complete Functional Module Manual

### 2.1 Operations Dashboard
- **8 Core KPI Tiles:**
  - `Total Sites`: Total active hoardings/unipoles/gantries with flagged review count.
  - `Active Sites`: Occupied displays currently generating campaign revenue.
  - `Vacant / Non-active`: Immediate inventory available for new sales proposals.
  - `Active Campaigns`: Live multi-site advertising bookings.
  - `Mounting Overdue`: Triggers when start date has arrived but physical mounting is unconfirmed.
  - `15-Day Validation Due`: Triggers when inspection photos with newspapers/GPS are required.
  - `Final Validation Due`: Triggers at campaign completion.
  - `Invoice Actions`: Flags pending GST invoice generation and courier dispatch.
- **Interactive Portfolio Occupancy Donut Chart:** Categorizes site performance into *Star (85%+)*, *Solid (65–84%)*, *Needs Attention (50–64%)*, and *Underperforming (<50%)*.
- **Live Action Alerts:** Prioritized task cards linking directly to the affected site and booking.

### 2.2 Sites Directory & Media Master
- **Attributes:** Site Code (`AMD-GT-001`, `AMD-HD-005`), Area, Full Address, Media Type (`Gantry`, `Hoarding`, `Unipole`, `Billboard`, `DOOH`), Lighting (`BL`, `FL`, `NL`), Size (W × H ft), Monthly Base Rate, GPS Coordinates, Google Maps Deep Link.
- **Multi-Image Management:** Upload and manage up to 12 site photos per display for automated presentations.

### 2.3 Automated PowerPoint Generator (PPTX)
- **16:9 Widescreen Presentation Standard:**
  - `Slide 1`: Uploadable full-bleed corporate introduction cover page.
  - `Slide 2`: Uploadable full-bleed company profile/credentials slide.
  - `Dynamic Site Slides`: 2-column split card (Left: yellow/navy brand specifications with site code, location, size, lighting, availability date, rate, GPS; Right: full-bleed site photograph).
  - `Final Slide`: Uploadable full-bleed thank-you / contact slide.
- **Bulk Filtering:** Filter sites by area, availability date threshold, and keyword before instant PPTX generation.

### 2.4 Proposal Builder & Commercial Letterhead
- **Left Control Panel:** Client selection/entry, campaign duration (days), validity (days), discount %, 18% GST tax rate, terms & conditions editor.
- **Selectable Picklist:** Set itemized Media Rate, Vendor Rate, and Printing Rate per site.
- **Right Live Letterhead Document:**
  - Media Buzz official logo and Ahmedabad OOH branding.
  - Campaign date range badge (e.g. `03 Sept 2026 to 01 Oct 2026 • 30 days`).
  - Itemized media plan breakdown table.
  - Commercial summary with Subtotal, Discount deduction, GST, and Grand Total.
  - Print-ready CSS for direct browser `Save as PDF` or hard-copy printing.

### 2.5 Electricity & Meter Tracker
- Tracks meter numbers, service numbers, T-numbers, utility providers (Torrent Power, UGVCL), billing months, units consumed, tariff rates, due dates, payment receipts, and overdue alerts.

### 2.6 Import / Export Engine
- **Excel Import (`.xlsx`, `.xls`):** Natively parses the Media Buzz standard format:
  `SR NO`, `AREA`, `LOCATION`, `MEDIA`, `LIGHT`, `W`, `H`, `SQ FT`, `AVAILABLITY`, `Selling Amount`, `Latitude Longitude`.
  - Automatically matches existing sites by Location/Area and updates rates, availability, dimensions, and GPS.
  - Automatically creates new sites with standardized codes (`AMD-GT-001`, `AMD-HD-008`).
- **Excel & CSV Export:** Generates clean `.xlsx` and `.csv` files matching the exact 11-column template.
- **JSON Database Snapshot:** Full database backup and one-click JSON database restore.

### 2.7 User Management & Role-Based Access Control (RBAC)
- **Admin:** Full workspace control, user creation, settings, and database management.
- **Manager:** Site management, campaign bookings, proposal creation, and performance reporting.
- **Staff:** Field operations, mounting status updates, meter logging, and photo validation uploads.
- **Viewer:** Read-only access to available inventory and proposal downloads.

---

## 3. Production Advantages & Strategic Value

| Feature / Dimension | Media Buzz Standalone Web App | Traditional WordPress Plugin | Enterprise SaaS Alternative |
|---|---|---|---|
| **Performance & Load Speed** | **Sub-50ms API response**, instant client-side transitions via React. | 800ms–2500ms due to WordPress PHP core and theme bloat. | Variable (dependent on multi-tenant cloud latency). |
| **Data Ownership & Privacy** | **100% Private Database** under your own MySQL server. | Private DB, but vulnerable to WP plugin security holes. | Third-party vendor stores your client lists & rates. |
| **Offline & Presentation Speed** | Instant client-side PPTX generation in under 2 seconds. | Slow server-side PHP memory limits during PPT export. | Watermarked or rate-limited export quotas. |
| **Security & Attack Surface** | Minimal attack surface (No WP core, no theme vulnerabilities, parameterized queries). | High risk (WP XML-RPC, brute-force admin logins, plugin conflicts). | Dependent on vendor security posture. |
| **Maintenance Cost** | **Zero subscription fees**. Runs on standard \$3–\$10/mo hosting (Hostinger/VPS). | Plugin license fees + WP maintenance overhead. | \$50–\$300 per user / month. |
| **Customization Flexibility** | Full source code access in React/Node.js for any custom feature. | Limited by WordPress hooks and PHP execution limits. | Rigid, closed-source proprietary software. |

---

## 4. Production Disadvantages, Trade-offs & Mitigations

### 4.1 Trade-offs & Disadvantages
1. **Self-Managed Infrastructure:**
   - *Risk:* If the Node.js process crashes or server reboots, the app requires a process manager (PM2 / Hostinger Node runner) to restart automatically.
   - *Mitigation:* System is configured with automatic restart scripts and Hostinger Node application manager.
2. **Database Backup Responsibility:**
   - *Risk:* Without WordPress backup plugins, manual or scheduled MySQL dumps must be configured.
   - *Mitigation:* Built-in **Export JSON Database Backup** in the Import/Export module allows one-click instant database snapshots from any browser.
3. **Email Delivery Configuration:**
   - *Risk:* Automated operational email reminders require valid SMTP credentials in `.env` rather than default PHP `mail()`.
   - *Mitigation:* Integrated standard SMTP environment parameters (`SMTP_HOST`, `SMTP_PORT`, `SMTP_USER`, `SMTP_PASS`).

---

## 5. Production Deployment & Hostinger Setup Checklist

### 5.1 Hostinger Node.js App Configuration
1. Log in to **Hostinger hPanel** $\rightarrow$ **Websites** $\rightarrow$ Select `mb.webtrionix.com`.
2. Navigate to **Node.js** application settings:
   - **Node.js Version:** `18.x` or `20.x`
   - **Application Root:** `/` (or root where repository is cloned)
   - **Application Startup File:** `server/src/index.js`
   - **Application URL:** `https://mb.webtrionix.com`

### 5.2 Environment Variables (`server/.env`)
Ensure the production `.env` file in the `server` directory contains:
```env
PORT=3000
NODE_ENV=production
DB_HOST=localhost
DB_PORT=3306
DB_USER=u245050038_mb_dashboard
DB_PASSWORD=YourStrongDatabasePasswordHere
DB_NAME=u245050038_mb_dashboard
JWT_SECRET=super_secure_jwt_secret_key_media_buzz_2026
ADMIN_EMAIL=admin@mediabuzz.com
ADMIN_PASSWORD=YourAdminPassword
```

### 5.3 Build & Deploy Commands
```bash
# 1. Install server dependencies
cd server && npm install --production

# 2. Build React client bundle
cd ../client && npm install && npm run build

# 3. Start or restart the Node.js application
# (On Hostinger, click "Restart Application" in the Node.js panel)
```

---

## 6. Security, Backups & Disaster Recovery

### 6.1 Database Backup Commands (Cron Job)
Set up a daily cron job on your server to export MySQL dumps:
```bash
# Daily automated backup at 02:00 AM
0 2 * * * mysqldump -u u245050038_mb_dashboard -p'YourPassword' u245050038_mb_dashboard > /home/backups/mb_backup_$(date +\%F).sql
```

### 6.2 Browser-Based One-Click Backup
- Navigate to **Import / Export** $\rightarrow$ Click **Export JSON Backup**.
- Saves the complete dataset (sites, campaigns, electricity bills, clients, vendors, invoices) to a portable `.json` file that can be restored instantly on any new server.

---

## 7. API & Data Reference

| Endpoint | Method | Role Required | Description |
|---|---|---|---|
| `/api/auth/login` | `POST` | Public | Authenticate user and receive JWT token. |
| `/api/dashboard` | `GET` | Authenticated | Fetch 8 KPI counts, alerts, and occupancy buckets. |
| `/api/sites` | `GET`, `POST` | Authenticated | List all inventory sites or create a new site. |
| `/api/sites/:id/images` | `POST` | Authenticated | Upload presentation photos for a site card. |
| `/api/sites/:id/images/:idx`| `DELETE`| Authenticated | Delete a specific site photo. |
| `/api/ppt-pages` | `GET` | Authenticated | Fetch URLs for Cover, Profile, and Ending presentation slides. |
| `/api/ppt-pages/:key` | `POST`, `DELETE`| Authenticated | Upload or remove fixed presentation slide images. |
| `/api/import/xlsx` | `POST` | Authenticated | Upload Excel workbook to batch-update sites and availability. |
| `/api/import/json` | `POST` | Authenticated | Restore complete database snapshot from JSON. |
| `/api/export/json` | `GET` | Authenticated | Export full database backup as JSON. |
| `/api/users` | `GET`, `POST` | Admin | List users or create new staff account with role. |
| `/api/users/:id` | `PUT`, `DELETE` | Admin | Update user details/password or delete user. |
| `/api/settings` | `GET`, `PUT` | Admin | Read or update company parameters and rules. |

---
*Generated for Media Buzz OOH Workspace • All Rights Reserved.*
