import { createHmac } from 'crypto';
import { SignJWT } from 'jose';

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ success: false, message: 'Method not allowed' });
  }

  const { email, code } = req.body;
  if (!email || !code || code.length !== 6) {
    return res.status(400).json({ success: false, message: '参数错误' });
  }

  const secret = process.env.JWT_SECRET;
  const cookies = parseCookies(req.headers.cookie || '');
  const storedHmac = cookies['_vc_hmac'];

  // 验证HMAC（检查当前和前一个时间窗口，防止边界情况）
  const timeWindow = Math.floor(Date.now() / 300000);
  let valid = false;
  for (let tw of [timeWindow, timeWindow - 1]) {
    const expectedHmac = createHmac('sha256', secret)
      .update(email.toLowerCase() + code + tw)
      .digest('hex');
    if (expectedHmac === storedHmac) {
      valid = true;
      break;
    }
  }

  if (!valid) {
    return res.status(400).json({ success: false, message: '验证码错误或已过期' });
  }

  // 签发JWT
  const secretKey = new TextEncoder().encode(secret);
  const token = await new SignJWT({ email: email.toLowerCase() })
    .setProtectedHeader({ alg: 'HS256' })
    .setExpirationTime('7d')
    .setIssuedAt()
    .sign(secretKey);

  // 设置Cookie
  const maxAge = 7 * 24 * 60 * 60; // 7天
  res.setHeader('Set-Cookie', [
    `_vc_token=${token}; Path=/; HttpOnly; SameSite=Strict; Max-Age=${maxAge}`,
    `_vc_hmac=; Path=/; HttpOnly; Max-Age=0` // 清除hmac cookie
  ]);

  return res.status(200).json({ success: true, message: '验证成功' });
}

function parseCookies(cookieStr) {
  const cookies = {};
  cookieStr.split(';').forEach(c => {
    const [key, ...val] = c.trim().split('=');
    if (key) cookies[key] = val.join('=');
  });
  return cookies;
}
