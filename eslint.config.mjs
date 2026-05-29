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
    rules: {
      // Semantic TS rules — formatting stays Prettier's job.
      // @typescript-eslint is already registered via office-addins/recommended.
      "@typescript-eslint/no-explicit-any": "error",
      "@typescript-eslint/consistent-type-imports": [
        "error",
        { prefer: "type-imports", fixStyle: "inline-type-imports" },
      ],
      "@typescript-eslint/no-non-null-assertion": "error",
      "@typescript-eslint/no-inferrable-types": "error",
    },
  },
  {
    ignores: ["dist/**", "node_modules/**", "coverage*/**"],
  },
];
