/** Jest runs the pure core (providers/llm) with no Office or browser needed. */
module.exports = {
  preset: "ts-jest",
  testEnvironment: "node",
  roots: ["<rootDir>/src"],
  testMatch: ["**/__tests__/**/*.test.ts"],
  // The core modules under test import nothing from Office/DOM.
  moduleFileExtensions: ["ts", "js", "json"],
};
