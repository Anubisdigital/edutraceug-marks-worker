# edutraceug-marks-worker

# Edutrace UG — Marks Worker (`edutraceug-marks-worker`)

Backend REST microservice for managing assessment score grids, Excel mark uploads, term result calculations (Percentage & CBC modes), and printable student report cards.

## Required Cloudflare Secrets

Before deploying, configure the following secrets in Cloudflare Workers using Wrangler or the Cloudflare Dashboard:

```bash
npx wrangler secret put ACCOUNT_SERVICE_FIREBASE
npx wrangler secret put CLOUDMERSIVE_API_KEY
npx wrangler secret put CLOUDINARY_URL
