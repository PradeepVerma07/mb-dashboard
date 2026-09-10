import mysql from 'mysql2/promise';
import dotenv from 'dotenv';
import path from 'path';
import {fileURLToPath} from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
dotenv.config();
dotenv.config({path: path.resolve(__dirname, '../../.env')});
dotenv.config({path: path.resolve(__dirname, '../.env')});
dotenv.config({path: path.resolve(process.cwd(), '.env')});

const rawHost = process.env.DB_HOST || 'localhost';
const host = (process.platform === 'win32' && (rawHost === 'localhost' || rawHost === '::1')) ? '127.0.0.1' : rawHost;

export const pool = mysql.createPool({
  host,
  port: Number(process.env.DB_PORT || 3306),
  user: process.env.DB_USER || 'root',
  password: process.env.DB_PASSWORD || '',
  database: process.env.DB_NAME || 'sitecontrol',
  socketPath: process.env.DB_SOCKET || undefined,
  waitForConnections: true,
  connectionLimit: 10,
  queueLimit: 100,
  connectTimeout: 10000,
  enableKeepAlive: true,
  keepAliveInitialDelay: 10000,
  decimalNumbers: true,
  charset: 'utf8mb4',
  ssl: String(process.env.DB_SSL).toLowerCase() === 'true' ? {} : undefined
});

export async function q(sql, params = []) {
  const sanitized = (params || []).map(p => p === undefined ? null : p);
  const [rows] = await pool.execute(sql, sanitized);
  return rows;
}

