# Use an official Node.js runtime as a parent image
FROM node:lts-slim as node-builder

# Set the working directory in the container
WORKDIR /usr/src/app

# Copy package.json and package-lock.json to the working directory
COPY package*.json ./

# Install Node.js and npm
RUN curl -fsSL https://deb.nodesource.com/setup_16.x | bash -

# Install Node.js dependencies
RUN npm install

# Copy the rest of the Node.js application code to the working directory
COPY . .

# Use an official Python runtime as a parent image
FROM python:3.9-slim

# Set the working directory in the container
WORKDIR /usr/src/app

# Copy installed Node.js modules from the previous stage
COPY --from=node-builder /usr/src/app /usr/src/app

# Install Puppeteer dependencies
RUN apt-get update && apt-get install -y \
    libnss3 \
    libxshmfence1 \
    libatk-bridge2.0-0 \
    libxcomposite1 \
    libxcursor1 \
    libxdamage1 \
    libxi6 \
    libxtst6 \
    libnss3 \
    libgconf-2-4 \
    libgbm-dev \
    libasound2 \
    libpangocairo-1.0-0 \
    libxrandr2 \
    libatk1.0-0 \
    ca-certificates \
    fonts-liberation \
    libappindicator3-1 \
    libgtk-3-0 \
    --no-install-recommends && \
    apt-get clean && \
    rm -rf /var/lib/apt/lists/*

# Copy Python requirements and install
COPY requirements.txt /app/
RUN pip install --no-cache-dir -r requirements.txt

# Copy all other application files
COPY . /app/

# Expose the port the app runs on
EXPOSE 8080

# Run Python application
CMD ["python", "bot.py"]
