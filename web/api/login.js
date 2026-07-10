import { SignJWT } from 'jose';
import { Redis } from '@upstash/redis';

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ success: false, message: 'Method not allowed' });
  }

  const { email, password } = req.body;
  if (!email || !password) {
    return res.status(400).json({ success: false, message: '请输入邮箱和密码' });
  }

  const redis = new Redis({
    url: process.env.UPSTASH_REDIS_REST_URL,
    token: process.env.UPSTASH_REDIS_REST_TOKEN,
  });

  // 从Redis读取用户（key: user:<email>，value: {password, name}）
  const userData = await redis.get(`user:${email.toLowerCase()}`);
  if (!userData || userData.password !== password) {
    await logAccess(redis, email.toLowerCase(), false, req, null);
    return res.status(400).json({ success: false, message: '邮箱或密码错误' });
  }

  // 签发JWT
  const secret = new TextEncoder().encode(process.env.JWT_SECRET);
  const token = await new SignJWT({
    email: email.toLowerCase(),
    name: userData.name || ''
  })
    .setProtectedHeader({ alg: 'HS256' })
    .setExpirationTime('7d')
    .setIssuedAt()
    .sign(secret);

  await logAccess(redis, email.toLowerCase(), true, req, userData);

  const maxAge = 7 * 24 * 60 * 60;
  res.setHeader('Set-Cookie', `_vc_token=${token}; Path=/; HttpOnly; SameSite=Strict; Max-Age=${maxAge}`);
  return res.status(200).json({ success: true, message: '登录成功' });
}

async function logAccess(redis, email, success, req, userData) {
  try {
    const entry = {
      email,
      name: userData?.name || '',
      org: userData?.org || '',
      success,
      time: new Date().toISOString(),
      ip: req.headers['x-forwarded-for'] || req.headers['x-real-ip'] || '',
      ua: req.headers['user-agent'] || '',
    };
    await redis.lpush('login_logs', JSON.stringify(entry));
    await redis.ltrim('login_logs', 0, 499);
  } catch (e) {
    console.error('Log error:', e);
  }
}
