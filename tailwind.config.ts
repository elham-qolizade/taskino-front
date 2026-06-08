import type { Config } from "tailwindcss";

const config: Config = {
  darkMode: "class",
  content: ["./src/**/*.{js,ts,jsx,tsx,mdx}"],
  theme: {
    extend: {
      fontFamily: {
        sans: ["var(--font-vazir)", "Tahoma", "Arial", "sans-serif"],
      },
      boxShadow: {
        soft: "0 16px 40px rgba(20, 31, 47, 0.08)",
      },
    },
  },
  plugins: [],
};

export default config;
