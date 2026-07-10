import { Resend } from 'resend';
import { createHmac } from 'crypto';

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ success: false, message: 'Method not allowed' });
  }

  const { email } = req.body;
  if (!email || !email.includes('@')) {
    return res.status(400).json({ success: false, message: '请输入有效的邮箱地址' });
  }

  const code = String(Math.floor(100000 + Math.random() * 900000));
  const secret = process.env.JWT_SECRET;
  const timeWindow = Math.floor(Date.now() / 300000); // 5分钟窗口
  const hmac = createHmac('sha256', secret)
    .update(email.toLowerCase() + code + timeWindow)
    .digest('hex');

  // 发送邮件
  const resend = new Resend(process.env.RESEND_API_KEY);
  try {
    await resend.emails.send({
      from: process.env.SENDER_EMAIL || 'noreply@example.com',
      to: email,
      subject: '中国乘用车海外数据看板 - 登录验证码',
      html: `
        <div style="font-family:'Microsoft YaHei',sans-serif;max-width:480px;margin:0 auto;padding:30px;background:#f8f9fa;border-radius:12px">
          <h2 style="color:#1a1a2e;margin-bottom:20px">中国乘用车海外数据看板</h2>
          <p style="color:#555;font-size:14px;margin-bottom:20px">您正在登录数据看板，验证码如下：</p>
          <div style="background:#1a1a2e;color:#C9A84C;font-size:32px;font-weight:700;letter-spacing:8px;text-align:center;padding:20px;border-radius:8px;margin-bottom:20px">${code}</div>
          <p style="color:#888;font-size:12px">验证码5分钟内有效，请勿泄露给他人。</p>
          <p style="color:#aaa;font-size:11px;margin-top:20px;border-top:1px solid #eee;padding-top:12px">广发证券发展研究中心 · 汽车行业研究小组</p>
        </div>
      `
    });
  } catch (e) {
    console.error('Email send error:', e);
    return res.status(500).json({ success: false, message: '邮件发送失败，请稍后重试' });
  }

  // 返回HMAC签名（前端不可见验证码，仅用于后续验证）
  res.setHeader('Set-Cookie', `_vc_hmac=${hmac}; Path=/; HttpOnly; SameSite=Strict; Max-Age=300`);
  return res.status(200).json({ success: true, message: '验证码已发送' });
}
