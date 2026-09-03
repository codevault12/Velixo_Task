/* eslint-disable no-undef */
module.exports = {
  preset: "ts-jest",
  testEnvironment: "jsdom",
  roots: ["<rootDir>/test"],
  testMatch: ["**/*.test.ts"],
  transform: {
    "^.+\\.tsx?$": [
      "ts-jest",
      {
        tsconfig: {
          module: "commonjs",
          target: "ES2020",
          esModuleInterop: true,
          lib: ["ES2020", "DOM", "WebWorker"],
          types: ["jest", "office-js", "office-runtime", "custom-functions-runtime"],
        },
      },
    ],
  },
};
