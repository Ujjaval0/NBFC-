import react from "@vitejs/plugin-react";
import { defineConfig } from "vite";

export default defineConfig({
  plugins: [react()],
  build: {
    chunkSizeWarningLimit: 1000,
    rollupOptions: {
      output: {
        manualChunks(id) {
          if (id.includes("node_modules")) {
            if (id.includes("recharts") || id.includes("d3")) {
              return "recharts";
            }
            if (id.includes("lucide-react")) {
              return "lucide";
            }
            return "vendor";
          }
        }
      }
    }
  }
});

