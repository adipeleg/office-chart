module.exports = {
  // setupFiles: ["<rootDir>/test/setup.ts"],
  roots: [
    "<rootDir>/lib",
    // "<rootDir>/client/src"
  ],
  transform: {
    "^.+\\.tsx?$": "ts-jest",
  },
  // setupFilesAfterEnv: [
  //   "./test/setup.ts",
  // // can have more setup files here
  // ],
  moduleFileExtensions: ["ts", "tsx", "js", "jsx", "json", "node"],
};