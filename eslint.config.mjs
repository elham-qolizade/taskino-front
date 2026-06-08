import next from "eslint-config-next";

const eslintConfig = [
  {
    ignores: ["Taskosaur/**", "taskino/**", ".next/**", "node_modules/**"],
  },
  ...next,
];

export default eslintConfig;
