import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  base: "./",
  plugins: [react()],
  server: { port: 5173, strictPort: true },
  build: {
    outDir: "dist",
    chunkSizeWarningLimit: 1600,
    rolldownOptions: {
      output: {
        codeSplitting: {
          groups: [
            {
              debugName: "app-dependencies",
              name(id: string) {
                if (
                  /node_modules[\\/](ag-grid-community|ag-grid-react)[\\/]/.test(
                    id,
                  )
                )
                  return "grid";
                if (id.includes("node_modules")) return "vendor";
                return null;
              },
            },
          ],
        },
      },
    },
  },
});
