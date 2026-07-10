import { jwtVerify } from 'jose';

export const config = {
  matcher: ['/dashboard.html', '/admin.html'],
};

export default async function middleware(req) {
  const cookieHeader = req.headers.get('cookie') || '';
  const cookies = Object.fromEntries(
    cookieHeader.split(';').map(c => {
      const [k, ...v] = c.trim().split('=');
      return [k, v.join('=')];
    })
  );
  const token = cookies['_vc_token'];

  if (!token) {
    return Response.redirect(new URL('/login.html', req.url));
  }

  try {
    const secret = new TextEncoder().encode(process.env.JWT_SECRET);
    const { payload } = await jwtVerify(token, secret);
    if (new URL(req.url).pathname === '/admin.html') {
      const admins = (process.env.ADMIN_EMAILS || '')
        .split(',')
        .map(email => email.trim().toLowerCase())
        .filter(Boolean);
      const email = String(payload.email || '').toLowerCase();
      if (admins.length === 0 || !admins.includes(email)) {
        return Response.redirect(new URL('/dashboard.html', req.url));
      }
    }
    return undefined; // pass through
  } catch (e) {
    // Token无效或过期，清除cookie并重定向
    const url = new URL('/login.html', req.url);
    return new Response(null, {
      status: 302,
      headers: {
        'Location': url.toString(),
        'Set-Cookie': '_vc_token=; Path=/; HttpOnly; Max-Age=0'
      }
    });
  }
}
