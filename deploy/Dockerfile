FROM node:20-alpine

WORKDIR /app

# Copy source
COPY package.json .
COPY src/ ./src/

# Install dependencies
RUN npm install --production

# Expose port
EXPOSE 3749

# Start
CMD ["node", "src/index.js"]
