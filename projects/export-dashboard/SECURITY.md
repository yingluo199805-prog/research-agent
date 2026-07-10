# Security gate for user management

The copied legacy implementation stores user passwords in clear text in Redis.
It must not be treated as production-ready on a new deployment until passwords
are migrated to a password-hash format and existing accounts have been reset or
migrated deliberately.

The copied code has been changed to fail closed for admin operations: an empty
`ADMIN_EMAILS` variable now denies access to the user list, login logs, and the
admin page. Configure a comma-separated administrator allowlist in Vercel
before deploying.

Required Vercel environment variables:

- `JWT_SECRET`
- `UPSTASH_REDIS_REST_URL`
- `UPSTASH_REDIS_REST_TOKEN`
- `ADMIN_EMAILS`

Optional email-login variables are `RESEND_API_KEY` and `SENDER_EMAIL`.
Never commit their values, `.vercel/`, `node_modules/`, local environment files,
or external data workbooks.
