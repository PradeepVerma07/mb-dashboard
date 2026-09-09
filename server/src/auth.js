import jwt from 'jsonwebtoken';

export function sign(user) {
  return jwt.sign(
    { id: user.id, email: user.email, name: user.name, role: (user.role || 'staff').toLowerCase() },
    process.env.JWT_SECRET,
    { expiresIn: process.env.JWT_EXPIRES_IN || '12h' }
  );
}

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

export function admin(req, res, next) {
  if ((req.user?.role || '').toLowerCase() !== 'admin') {
    return res.status(403).json({ message: 'Administrator access required' });
  }
  next();
}

export function managerOrAdmin(req, res, next) {
  const role = (req.user?.role || '').toLowerCase();
  if (role !== 'admin' && role !== 'manager') {
    return res.status(403).json({ message: 'Manager or Administrator access required' });
  }
  next();
}

export function notViewer(req, res, next) {
  const role = (req.user?.role || '').toLowerCase();
  if (role === 'viewer') {
    return res.status(403).json({ message: 'Viewer role has read-only access. Cannot create, edit, or delete records.' });
  }
  next();
}
