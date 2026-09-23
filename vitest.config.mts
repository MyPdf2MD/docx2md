import { defineConfig } from "vitest/config";

/**
 * `node`, with no DOM anywhere. That is the point: the parser used to need
 * jsdom for `DOMParser`, and the bundled XML reader in `src/xml.ts` exists so
 * that the code under test is the code that ships, in every runtime.
 */
export default defineConfig({
  test: {
    environment: "node",
    include: ["test/**/*.test.ts"],
    coverage: {
      provider: "v8",
      include: ["src/**/*.ts"],
      // Test scaffolding: builders for the synthetic .docx fixtures, not
      // shipped behaviour.
      exclude: ["src/cli.ts"],
    },
  },
});
