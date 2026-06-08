# syntax=docker/dockerfile:1

FROM node:20-alpine AS deps
WORKDIR /app

# Nur Manifeste kopieren, damit der Layer-Cache bei reinen Code-Änderungen greift
COPY package.json package-lock.json ./
RUN npm ci --omit=dev

FROM node:20-alpine AS runtime
WORKDIR /app

ENV NODE_ENV=production

# Produktions-Abhängigkeiten aus dem deps-Stage übernehmen
COPY --from=deps /app/node_modules ./node_modules

# Anwendungscode kopieren (siehe .dockerignore für Ausschlüsse)
COPY . .

# Als non-root laufen (Image bringt den User "node" mit)
USER node

# MCP-Server kommuniziert über stdio
ENTRYPOINT ["node", "index.js"]
