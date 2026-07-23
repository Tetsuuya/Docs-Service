FROM node:20-alpine

# Set working directory
WORKDIR /app

# Copy package management files
COPY package*.json ./

# Install production dependencies
RUN npm ci --only=production

# Copy source code
COPY . .

# Expose application port
EXPOSE 3000

# Start Docs-Service server
CMD ["node", "server.js"]
