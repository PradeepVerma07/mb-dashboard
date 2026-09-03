import mysql from 'mysql2/promise';
import dotenv from 'dotenv';
import path from 'path';
import {fileURLToPath} from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
dotenv.config();
dotenv.config({path: path.resolve(__dirname, '../../.env')});
dotenv.config({path: path.resolve(__dirname, '../.env')});
dotenv.config({path: path.resolve(process.cwd(), '.env')});

export const pool = mysql.createPool({
  host: process.env.DB_HOST || '127.0.0.1',
  port: Number(process.env.DB_PORT || 3306),
  user: process.env.DB_USER || 'root',
  password: process.env.DB_PASSWORD || '',
  database: process.env.DB_NAME || 'sitecontrol',
  waitForConnections: true,
  connectionLimit: 10,
  queueLimit: 0,
  decimalNumbers: true,
  ssl: String(process.env.DB_SSL).toLowerCase() === 'true' ? {} : undefined
});

export async function q(sql, params = []) {
  const [rows] = await pool.execute(sql, params);
  return rows;
}

