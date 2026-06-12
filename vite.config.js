import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// Static SPA — no backend. The 114 surah JSON files live in public/data and are
// fetched at runtime; the Word file is built entirely in the browser (see src/buildDocx.js).
export default defineConfig({
  // Relative base so the build works whether it's served at the domain root
  // (quran-similarities.com/) or under a project subpath
  // (raghad-mezied.github.io/quran-similarities/).
  base: "./",
  plugins: [react()],
});
