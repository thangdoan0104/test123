/** @type {import('next').NextConfig} */
const repositoryName = process.env.GITHUB_REPOSITORY?.split("/")[1] ?? "";
const isGithubPagesBuild = process.env.GITHUB_PAGES === "true";
const isUserOrOrgPage = repositoryName.endsWith(".github.io");
const githubPagesBasePath =
  isGithubPagesBuild && repositoryName && !isUserOrOrgPage ? `/${repositoryName}` : "";

const nextConfig = {
  reactStrictMode: true,

  // GitHub Pages only serves static files. `next build` writes the site to /out.
  output: "export",
  trailingSlash: true,

  // Required for static export when using next/image later.
  images: {
    unoptimized: true
  },

  // Required for project pages such as https://username.github.io/repo-name/
  // Local dev and Vercel deploys keep this empty.
  basePath: githubPagesBasePath,
  assetPrefix: githubPagesBasePath ? `${githubPagesBasePath}/` : undefined,

  webpack: (config) => {
    // Fabric 5 is browser-only in this app, but its npm package can reference
    // Node canvas/jsdom during dependency resolution. These aliases prevent
    // Next's static build from trying to bundle native Node modules for the browser.
    config.resolve.alias = {
      ...(config.resolve.alias || {}),
      canvas: false,
      jsdom: false
    };

    config.resolve.fallback = {
      ...(config.resolve.fallback || {}),
      fs: false,
      path: false,
      stream: false,
      util: false,
      buffer: false
    };

    return config;
  }
};

export default nextConfig;
