import type { NextConfig } from "next";

const isGitHubPages = process.env.GITHUB_PAGES === "true";
const repositoryName = process.env.GITHUB_REPOSITORY?.split("/")[1] ?? "rota-facil";

const nextConfig: NextConfig = isGitHubPages
  ? {
      output: "export",
      basePath: `/${repositoryName}`,
      trailingSlash: true,
      images: { unoptimized: true },
      typescript: { tsconfigPath: "tsconfig.pages.json" },
    }
  : {};

export default nextConfig;
