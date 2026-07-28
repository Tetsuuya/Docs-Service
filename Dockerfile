FROM node:20-alpine

# Install Python and system dependencies for PPTX processing and canvas
RUN apk add --no-cache \
    python3 \
    py3-pip \
    py3-lxml \
    # Canvas native dependencies
    build-base \
    g++ \
    cairo-dev \
    jpeg-dev \
    pango-dev \
    giflib-dev \
    pixman-dev \
    pangomm-dev \
    libjpeg-turbo-dev \
    freetype-dev \
    pkgconfig \
    && pip3 install --break-system-packages --no-cache-dir python-pptx

# Set working directory
WORKDIR /app

# Copy package management files
COPY package*.json ./

# Install production dependencies
RUN npm ci --omit=dev

# Copy source code
COPY . .

# Expose application port
EXPOSE 3000

# Start Docs-Service server
CMD ["node", "server.js"]
