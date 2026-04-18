import type { Config } from 'tailwindcss';

const config: Config = {
  content: ["./src/**/*.{js,ts,jsx,tsx,mdx}"],
  theme: {
    extend: {
      colors: {
        ink: "#07111f",
        paper: "#f6f3ec",
        glow: "#f7c76a",
        accent: "#5d7cff",
        mint: "#6ad5b6"
      },
      boxShadow: {
        soft: "0 24px 80px rgba(7, 17, 31, 0.18)"
      },
      backgroundImage: {
        "hero-grid": "radial-gradient(circle at 20% 20%, rgba(109, 125, 255, 0.20), transparent 30%), radial-gradient(circle at 80% 0%, rgba(255, 190, 96, 0.22), transparent 28%), linear-gradient(180deg, #101b32 0%, #07111f 100%)"
      }
    }
  },
  plugins: []
};

export default config;
