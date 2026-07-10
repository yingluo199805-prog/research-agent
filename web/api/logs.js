import { jwtVerify } from 'jose';
import { Redis } from '@upstash/redis';

export default async function handler(req, res) {
  if (req.method !== 'GET') {
    return res.status(405).json({ success: false, message: 'Method not allowed' });
  }

  // 验证JWT（只有登录用户可以查看日志）
  const cookies = parseCookies(req.headers.cookie || '');
  const token = cookies['_vc_token'];
  if (!token) {
    return res.status(401).json({ success: false, message: '未登录' });
  }

  try {
    const secret = new TextEncoder().encode(process.env.JWT_SECRET);
    const { payload } = await jwtVerify(token, secret);

    // 可选：限制只有管理员可以查看（检查是否在ADMIN_EMAILS中）
    const admins = (process.env.ADMIN_EMAILS || '')
      .split(',')
      .map(e => e.trim().toLowerCase())
      .filter(Boolean);
    if (admins.length === 0 || !admins.includes(String(payload.email || '').toLowerCase())) {
      return res.status(403).json({ success: false, message: '无权限' });
    }
  } catch (e) {
    return res.status(401).json({ success: false, message: 'Token无效' });
  }

  // 读取日志
  try {
    const redis = new Redis({
      url: process.env.UPSTASH_REDIS_REST_URL,
      token: process.env.UPSTASH_REDIS_REST_TOKEN,
    });
    const count = parseInt(req.query.count || '50', 10);
    const logs = await redis.lrange('login_logs', 0, Math.min(count, 500) - 1);
    const parsed = logs.map(l => typeof l === 'string' ? JSON.parse(l) : l);
    return res.status(200).json({ success: true, logs: parsed });
  } catch (e) {
    console.error('Logs read error:', e);
    return res.status(500).json({ success: false, message: '读取日志失败' });
  }
}

function parseCookies(cookieStr) {
  const cookies = {};
  cookieStr.split(';').forEach(c => {
    const [key, ...val] = c.trim().split('=');
    if (key) cookies[key] = val.join('=');
  });
  return cookies;
}
