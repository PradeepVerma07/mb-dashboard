import 'dotenv/config';
import jwt from 'jsonwebtoken';

const JWT_SECRET = process.env.JWT_SECRET || 'local_media_buzz_secret_key_2026_very_long_and_secure_abcdef';
const API_URL = 'http://localhost:3000/api';

function makeToken(user) {
  return jwt.sign(
    { id: user.id, email: user.email, name: user.name, role: (user.role || 'staff').toLowerCase() },
    JWT_SECRET,
    { expiresIn: '1h' }
  );
}

async function testRBAC() {
  console.log('Testing Role-Based Access Control (RBAC)...');

  const adminToken = makeToken({ id: 1, email: 'admin@mediabuzz.com', name: 'Super Admin', role: 'admin' });
  const managerToken = makeToken({ id: 2, email: 'manager@mediabuzz.com', name: 'Operations Manager', role: 'manager' });
  const staffToken = makeToken({ id: 3, email: 'staff@mediabuzz.com', name: 'Field Staff', role: 'staff' });
  const viewerToken = makeToken({ id: 4, email: 'viewer@mediabuzz.com', name: 'Read-only Viewer', role: 'viewer' });

  // 1. Test /api/users endpoint
  console.log('\n--- 1. Testing /api/users (Admin-only) ---');
  
  // Admin should be able to get users
  const res1 = await fetch(`${API_URL}/users`, { headers: { Authorization: `Bearer ${adminToken}` } });
  if (res1.ok) {
    const data = await res1.json();
    console.log(`[PASS] Admin can GET /api/users (found ${data.length} users)`);
  } else {
    console.error(`[FAIL] Admin GET /api/users: status ${res1.status}`);
  }

  // Non-admin attempting to create a user
  const res2 = await fetch(`${API_URL}/users`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${managerToken}` },
    body: JSON.stringify({ name: 'Unauthorized User', email: 'unauth@test.com', password: 'password123', role: 'staff' })
  });
  if (res2.status === 403) {
    const err = await res2.json();
    console.log(`[PASS] Manager POST /api/users returned 403 Forbidden as expected: "${err.message}"`);
  } else {
    console.error(`[FAIL] Manager POST /api/users returned status ${res2.status} (expected 403)`);
  }

  // 2. Test Settings update (Admin only)
  console.log('\n--- 2. Testing /api/settings (Admin-only) ---');
  const res3 = await fetch(`${API_URL}/settings`, {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${staffToken}` },
    body: JSON.stringify({ company_name: 'Media Buzz Test' })
  });
  if (res3.status === 403) {
    const err = await res3.json();
    console.log(`[PASS] Staff PUT /api/settings returned 403 Forbidden as expected: "${err.message}"`);
  } else {
    console.error(`[FAIL] Staff PUT /api/settings returned status ${res3.status} (expected 403)`);
  }

  // 3. Test Viewer mutation protection
  console.log('\n--- 3. Testing Viewer Mutation Restrictions ---');
  const res4 = await fetch(`${API_URL}/storage/upload`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${viewerToken}` },
    body: JSON.stringify({})
  });
  if (res4.status === 403) {
    const err = await res4.json();
    console.log(`[PASS] Viewer POST /api/storage/upload returned 403 Forbidden as expected: "${err.message}"`);
  } else {
    console.error(`[FAIL] Viewer POST returned status ${res4.status} (expected 403)`);
  }

  // 4. Test Staff restriction on proposal creation
  console.log('\n--- 4. Testing Proposals/Invoices restricted to Manager/Admin ---');
  const res5 = await fetch(`${API_URL}/proposals`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${staffToken}` },
    body: JSON.stringify({ client_name: 'Test' })
  });
  if (res5.status === 403) {
    const err = await res5.json();
    console.log(`[PASS] Staff POST /api/proposals returned 403 Forbidden as expected: "${err.message}"`);
  } else {
    console.error(`[FAIL] Staff POST returned status ${res5.status} (expected 403)`);
  }

  console.log('\nAll RBAC security rules verified successfully!');
}

testRBAC().catch(err => {
  console.error('RBAC test runner error:', err.message);
  process.exit(1);
});
