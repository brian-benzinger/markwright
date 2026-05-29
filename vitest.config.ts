import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    coverage: {
      provider: "v8",
      include: ["src/convert/**"],
      reporter: ["text", "html"],
      thresholds: { lines: 90, statements: 90, functions: 100, branches: 80 },
    },
  },
});
