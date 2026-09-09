# Authentication & Role-Based Access Control (RBAC) Documentation
### Media Buzz OOH Operations Management System

---

## 1. Overview & Architecture

The Media Buzz Out-of-Home (OOH) Operations Platform uses a stateless, token-based authentication architecture with **Role-Based Access Control (RBAC)** implemented end-to-end across both the Node.js / Express backend and the React 18 single-page application.

```
       ┌───────────────────────────────┐
       │   React 18 Frontend Client    │
       │  (Token stored in sc_token)   │
       └──────────────┬────────────────┘
                      │ HTTP Request + Header:
                      │ Authorization: Bearer <JWT>
                      ▼
       ┌───────────────────────────────┐
       │     Express.js API Router     │
       └──────────────┬────────────────┘
                      │
                      ▼
       ┌───────────────────────────────┐
       │   1. auth(req, res, next)     │ ──▶ 401 Unauthorized (if invalid / expired)
       │      Validates JWT Signature  │
       │      Attaches req.user        │
       └──────────────┬────────────────┘
                      │
                      ▼
       ┌───────────────────────────────┐
       │   2. Role Guard Middleware    │ ──▶ 403 Forbidden (if role insufficient)
       │      requireRole(...roles)    │
       │      admin / managerOrAdmin   │
       │      notViewer                │
       └──────────────┬────────────────┘
                      │
                      ▼
       ┌───────────────────────────────┐
       │   3. Controller / Route Logic │
       │      Database Queries / CRUD  │
       └───────────────────────────────┘
```

---

## 2. Deep Dive: `server/src/auth.js`

The file [`server/src/auth.js`](file:///c:/Users/reeha/OneDrive/Desktop/MB%20DASH/server/src/auth.js) is the core security module of the backend. It exports helper functions and Express middleware for token generation and access authorization.

### 2.1. `sign(user)` — Token Issuance
```javascript
export function sign(user) {
  return jwt.sign(
    { 
      id: user.id, 
      email: user.email, 
      name: user.name, 
      role: (user.role || 'staff').toLowerCase() 
    },
    process.env.JWT_SECRET,
    { expiresIn: process.env.JWT_EXPIRES_IN || '12h' }
  );
}
```
- **Payload Structure**:
  - `id` *(number)*: Unique user identifier in the `users` table.
  - `email` *(string)*: User login email.
  - `name` *(string)*: User display name.
  - `role` *(string)*: Security role, automatically normalized to lowercase (`admin`, `manager`, `staff`, or `viewer`).
- **Signature**: Signed using HMAC-SHA256 with the secret key configured in `process.env.JWT_SECRET`.
- **Expiration**: Defaults to `12h` (configurable via `process.env.JWT_EXPIRES_IN`).

---

### 2.2. `auth(req, res, next)` — Session Authentication Middleware
```javascript
export function auth(req, res, next) {
  const h = req.headers.authorization || '';
  const token = h.startsWith('Bearer ') ? h.slice(7) : '';
  if (!token) return res.status(401).json({ message: 'Authentication required' });
  try {
    req.user = jwt.verify(token, process.env.JWT_SECRET);
    if (req.user && req.user.role) req.user.role = req.user.role.toLowerCase();
    next();
  } catch {
    return res.status(401).json({ message: 'Session expired' });
  }
}
```
- **Header Format**: Expects an `Authorization` HTTP header with `Bearer <token>`.
- **Verification**: Verifies cryptographic integrity and validity timestamp against `JWT_SECRET`.
- **Context Injection**: Attaches the decoded payload to `req.user` with a lowercase `role` property.
- **Error Status**:
  - Missing token $\rightarrow$ `HTTP 401 Unauthorized` (`{ message: "Authentication required" }`).
  - Expired / tampered token $\rightarrow$ `HTTP 401 Unauthorized` (`{ message: "Session expired" }`).

---

### 2.3. `requireRole(...roles)` — Dynamic Role Authorization Middleware
```javascript
export function requireRole(...roles) {
  const normalized = roles.map(r => r.toLowerCase());
  return (req, res, next) => {
    const userRole = (req.user?.role || 'staff').toLowerCase();
    if (!normalized.includes(userRole)) {
      return res.status(403).json({
        message: `Access denied. Requires one of: ${roles.join(', ')}. Current role: ${userRole}`
      });
    }
    next();
  };
}
```
- Accepts one or more permitted roles (e.g. `requireRole('admin', 'manager')`).
- Evaluates `req.user.role` case-insensitively.
- If unauthorized, immediately halts execution with `HTTP 403 Forbidden`.

---

### 2.4. Specialized Convenience Role Guards

To ensure concise, readable, and consistent route definitions across [`server/src/index.js`](file:///c:/Users/reeha/OneDrive/Desktop/MB%20DASH/server/src/index.js), pre-configured middleware guards are provided:

```javascript
// 1. Strict Administrator Only
export function admin(req, res, next) {
  if ((req.user?.role || '').toLowerCase() !== 'admin') {
    return res.status(403).json({ message: 'Administrator access required' });
  }
  next();
}

// 2. Operational Leads (Manager or Administrator)
export function managerOrAdmin(req, res, next) {
  const role = (req.user?.role || '').toLowerCase();
  if (role !== 'admin' && role !== 'manager') {
    return res.status(403).json({ message: 'Manager or Administrator access required' });
  }
  next();
}

// 3. Mutation Blocker for Viewers
export function notViewer(req, res, next) {
  const role = (req.user?.role || '').toLowerCase();
  if (role === 'viewer') {
    return res.status(403).json({ message: 'Viewer role has read-only access. Cannot create, edit, or delete records.' });
  }
  next();
}
```

---

## 3. Workspace Security Roles & Capabilities Matrix

The workspace defines 4 explicit operational security roles:

| Capability / Module | 👑 Admin (`admin`) | 💼 Manager (`manager`) | 🛠️ Staff (`staff`) | 👁️ Viewer (`viewer`) |
| :--- | :---: | :---: | :---: | :---: |
| **Browse Dashboard & Performance Reports** | ✓ | ✓ | ✓ | ✓ |
| **Download Generated PPT Decks & Excel Exports** | ✓ | ✓ | ✓ | ✓ |
| **View Storage Vault & Occupancy Archives** | ✓ | ✓ | ✓ | ✓ |
| **View Sites Inventory Directory** | ✓ | ✓ | ✓ | ✓ |
| **View Campaigns & Operations Tracker** | ✓ | ✓ | ✓ | ✓ |
| **Update Campaign Mounting & Printing Status** | ✓ | ✓ | ✓ | ✕ |
| **Record & Pay Electricity Meter Bills** | ✓ | ✓ | ✓ | ✕ |
| **Create & Edit Sites Inventory** | ✓ | ✓ | ✕ | ✕ |
| **Create & Edit Campaign Bookings** | ✓ | ✓ | ✕ | ✕ |
| **Create Client Proposals & Financial Invoices** | ✓ | ✓ | ✕ | ✕ |
| **Upload Files to Storage & Take Occupancy Snapshots** | ✓ | ✓ | ✕ | ✕ |
| **Batch Delete Records & Storage Vault Archives** | ✓ | ✓ | ✕ | ✕ |
| **Single Record Deletions (Sites, Campaigns, Invoices)** | ✓ | ✓ | ✕ | ✕ |
| **Import CSV / Excel Inventory** | ✓ | ✓ | ✕ | ✕ |
| **Access User Management (Add, Edit, Delete Users)** | ✓ | ✕ | ✕ | ✕ |
| **Modify Workspace Settings & Retention Policies** | ✓ | ✕ | ✕ | ✕ |

---

## 4. API Route Protection Map (`server/src/index.js`)

Below is the mapping of backend endpoints to their respective security requirements:

| Endpoint | Method | Middleware Guard | Authorized Roles | Description |
| :--- | :---: | :--- | :--- | :--- |
| `/api/auth/login` | `POST` | *Public* | Anyone | Authenticate with email & password, returns JWT token |
| `/api/auth/me` | `GET` | `auth` | All authenticated | Retrieve currently authenticated user profile |
| `/api/users` | `GET` | `auth, admin` | `admin` | List all workspace users |
| `/api/users` | `POST` | `auth, admin` | `admin` | Create new user and assign role |
| `/api/users/:id` | `PUT` | `auth, admin` | `admin` | Update user credentials or role |
| `/api/users/:id` | `DELETE` | `auth, admin` | `admin` | Delete workspace user account |
| `/api/settings` | `GET` | `auth` | All authenticated | Read workspace configuration |
| `/api/settings` | `PUT` | `auth, admin` | `admin` | Update operational rules and company defaults |
| `/api/storage` | `GET` | `auth` | All authenticated | Browse storage vault files |
| `/api/storage/upload` | `POST` | `auth, notViewer` | `admin`, `manager`, `staff` | Upload document to storage |
| `/api/storage/snapshot-occupancy` | `POST` | `auth, notViewer` | `admin`, `manager`, `staff` | Archive monthly occupancy snapshot |
| `/api/storage/batch-delete` | `POST` | `auth, managerOrAdmin` | `admin`, `manager` | Batch delete files from storage vault |
| `/api/storage/:id` | `DELETE` | `auth, managerOrAdmin` | `admin`, `manager` | Delete a single storage vault file |
| `/api/import/:entity` | `POST` | `auth, managerOrAdmin` | `admin`, `manager` | Import Excel/CSV data |
| `/api/proposals` | `POST` | `auth, managerOrAdmin` | `admin`, `manager` | Create commercial proposal |
| `/api/invoices` | `POST` | `auth, managerOrAdmin` | `admin`, `manager` | Create client invoice |
| `/api/:entity` | `POST` | `auth, notViewer` | `admin`, `manager`, `staff`* | Create new record (*proposals/invoices restricted to mgr/admin) |
| `/api/:entity/:id` | `PUT` | `auth, notViewer` | `admin`, `manager`, `staff`* | Edit record |
| `/api/:entity/:id` | `DELETE` | `auth, managerOrAdmin` | `admin`, `manager` | Delete / archive record |
| `/api/:entity/batch-delete` | `POST` | `auth, managerOrAdmin` | `admin`, `manager` | Delete selected records in bulk |

---

## 5. Client-Side Implementation & UI Protection (`client/src/main.jsx`)

The client application enforces user roles at three levels:
1. **Client Route Guards**: Non-admins attempting to access `/#/settings` are automatically redirected to `/#/dashboard`. Staff and Viewers navigating to `/#/proposals` or `/#/ppt` are similarly routed back.
2. **Sidebar Navigation Filtering**: Only routes permitted for the user's role are rendered in the navigation menu (`visibleNavGroups`).
3. **Action & Modal Controls**:
   - **User Management Tab** (`/#/settings`): If a non-admin reaches the tab, an administrative security shield is displayed:
     > 🛡️ **Administrator Privileges Required**
     > User account provisioning, security role assignment, and access control management are restricted strictly to Workspace Administrators.
   - **Role Metric Counters**: Displays real-time counts for Total Members, Admins, Managers, Staff, and Viewers.
   - **Self-Lockout Protection**: Delete button is disabled on the currently logged-in account to prevent accidental lockout.
   - **Dynamic Role Preview**: When adding or editing users, changing the role in the dropdown dynamically renders a capabilities card detailing allowed operations and restrictions.
   - **Batch Delete & Delete Buttons**: Hidden for Staff and Viewers.
   - **Viewer Mode Banners**: When logged in as Viewer, a topbar badge (`🔒 Read-Only Workspace`) and banner notifications inform the user that changes cannot be made.

---

## 6. Automated Testing & Verification

A dedicated test suite [`server/src/test_rbac_e2e.js`](file:///c:/Users/reeha/OneDrive/Desktop/MB%20DASH/server/src/test_rbac_e2e.js) is provided to verify role boundaries.

### Running the Tests
```bash
# In server directory:
node src/test_rbac_e2e.js
```

### Test Output
```
Testing Role-Based Access Control (RBAC)...

--- 1. Testing /api/users (Admin-only) ---
[PASS] Admin can GET /api/users (found 1 users)
[PASS] Manager POST /api/users returned 403 Forbidden as expected: "Administrator access required"

--- 2. Testing /api/settings (Admin-only) ---
[PASS] Staff PUT /api/settings returned 403 Forbidden as expected: "Administrator access required"

--- 3. Testing Viewer Mutation Restrictions ---
[PASS] Viewer POST /api/storage/upload returned 403 Forbidden as expected: "Viewer role has read-only access. Cannot create, edit, or delete records."

--- 4. Testing Proposals/Invoices restricted to Manager/Admin ---
[PASS] Staff POST /api/proposals returned 403 Forbidden as expected: "Manager or Administrator access required"

All RBAC security rules verified successfully!
```

---

## 7. Configuration Variables (`.env`)

The authentication layer relies on the following environment variables:

| Variable | Default Value | Description |
| :--- | :--- | :--- |
| `JWT_SECRET` | *required* | Strong cryptographic secret used to sign and verify JWT tokens |
| `JWT_EXPIRES_IN` | `12h` | Lifetime duration before a token expires |
| `ADMIN_EMAIL` | `admin@mediabuzz.com` | Default seed administrator account |
| `ADMIN_PASSWORD` | `admin123` | Initial administrator login password |
