import officeAddins from "eslint-plugin-office-addins";
import tsParser from "@typescript-eslint/parser";
import globals from "globals";

export default [
  ...officeAddins.configs.recommended,
  {
    plugins: {
      "office-addins": officeAddins,
    },
    languageOptions: {
      parser: tsParser,
      globals: {
        ...globals.browser,
        ...globals.node,
        Office: "readonly",
        OfficeRuntime: "readonly",
        Word: "readonly",
      },
    },
  },
  {
    ignores: ["dist/**", "node_modules/**"],
  },
];
