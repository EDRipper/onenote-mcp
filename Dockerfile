FROM node:22-bookworm-slim AS build
WORKDIR /app
COPY package*.json tsconfig.json ./
RUN npm ci
COPY src ./src
RUN npm run build && npm prune --omit=dev

FROM node:22-bookworm-slim
# Chromium renders Mermaid diagrams; DejaVu fonts are used for SVG and Mermaid text.
RUN apt-get update \
 && apt-get install -y --no-install-recommends chromium fonts-dejavu-core wget \
 && rm -rf /var/lib/apt/lists/*
WORKDIR /app
ENV NODE_ENV=production PORT=3000 CHROMIUM_PATH=/usr/bin/chromium
COPY --from=build /app/node_modules ./node_modules
COPY --from=build /app/dist ./dist
COPY package.json ./
USER node
EXPOSE 3000
HEALTHCHECK CMD wget -qO- http://127.0.0.1:3000/health || exit 1
CMD ["node", "dist/index.js"]
