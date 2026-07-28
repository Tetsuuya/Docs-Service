FROM node:20-alpine

# Install Python and dependencies for PPTX processing
RUN apk add --no-cache \
    python3 \
    py3-pip \
    py3-lxml \
    && pip3 install --no-cache-dir python-pptx

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
