import { jwtVerify } from 'jose';
import { Redis } from '@upstash/redis';

export default async function handler(req, res) {
  // 验证管理员身份
  const admin = await verifyAdmin(req);
  if (!admin) {
    return res.status(401).json({ success: false, message: '无权限' });
  }

  const redis = new Redis({
    url: process.env.UPSTASH_REDIS_REST_URL,
    token: process.env.UPSTASH_REDIS_REST_TOKEN,
  });

  if (req.method === 'GET') {
    // 获取所有用户
    const keys = await redis.keys('user:*');
    const users = [];
    for (const key of keys) {
      const data = await redis.get(key);
      const email = key.replace('user:', '');
      users.push({ email, name: data.name || '', org: data.org || '', createdAt: data.createdAt || '' });
    }
    users.sort((a, b) => a.email.localeCompare(b.email));
    return res.status(200).json({ success: true, users, count: users.length });
  }

  if (req.method === 'POST') {
    // 批量添加用户：[{email, password, name}]
    const { users } = req.body;
    if (!Array.isArray(users) || users.length === 0) {
      return res.status(400).json({ success: false, message: '请提供用户列表' });
    }
    let added = 0;
    let skipped = 0;
    for (const u of users) {
      if (!u.email || !u.password) continue;
      const key = `user:${u.email.toLowerCase()}`;
      const exists = await redis.exists(key);
      if (exists) { skipped++; continue; }
      await redis.set(key, {
        password: u.password,
        name: u.name || '',
        org: u.org || '',
        createdAt: new Date().toISOString()
      });
      added++;
    }
    return res.status(200).json({ success: true, added, skipped, message: `添加${added}人，跳过${skipped}人（已存在）` });
  }

  if (req.method === 'DELETE') {
    // 删除用户：{email}
    const { email } = req.body;
    if (!email) {
      return res.status(400).json({ success: false, message: '请提供邮箱' });
    }
    await redis.del(`user:${email.toLowerCase()}`);
    return res.status(200).json({ success: true, message: `已删除 ${email}` });
  }

  return res.status(405).json({ success: false, message: 'Method not allowed' });
}

async function verifyAdmin(req) {
  try {
    const cookies = parseCookies(req.headers.cookie || '');
    const token = cookies['_vc_token'];
    if (!token) return null;
    const secret = new TextEncoder().encode(process.env.JWT_SECRET);
    const { payload } = await jwtVerify(token, secret);
    const admins = (process.env.ADMIN_EMAILS || '')
      .split(',')
      .map(e => e.trim().toLowerCase())
      .filter(Boolean);
    if (admins.length === 0 || !admins.includes(String(payload.email || '').toLowerCase())) return null;
    return payload;
  } catch (e) { return null; }
}

function parseCookies(cookieStr) {
  const cookies = {};
  cookieStr.split(';').forEach(c => {
    const [key, ...val] = c.trim().split('=');
    if (key) cookies[key] = val.join('=');
  });
  return cookies;
}
