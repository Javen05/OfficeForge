/** @type {import('next').NextConfig} */
const isGithubActions = process.env.GITHUB_ACTIONS === 'true';
const repository = process.env.GITHUB_REPOSITORY?.split('/')[1] ?? '';
const basePath = isGithubActions && repository ? `/${repository}` : '';

const nextConfig = {
  typedRoutes: true,
  output: 'export',
  trailingSlash: true,
  basePath,
  assetPrefix: basePath,
  images: {
    unoptimized: true
  }
};

export default nextConfig;
