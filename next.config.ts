import type { NextConfig } from "next";
import path from "node:path";

// Backend origin to proxy /api requests to. Override via env on Vercel if it changes.
const backendOrigin = process.env.BACKEND_ORIGIN ?? "http://46.249.98.55:3000";

const nextConfig: NextConfig = {
  reactStrictMode: true,
  turbopack: {
    root: path.resolve(__dirname),
  },
  // Proxy /api/* to the backend server-side so the HTTPS frontend can talk to the
  // HTTP backend without mixed-content/CORS issues. Set NEXT_PUBLIC_API_URL=/api.
  async rewrites() {
    return [
      {
        source: "/api/:path*",
        destination: `${backendOrigin}/api/:path*`,
      },
    ];
  },
};

export default nextConfig;
