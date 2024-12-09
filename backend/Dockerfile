# Use a newer Node.js image based on a supported Debian version
FROM node:16

# Set the working directory inside the container
WORKDIR /app

# Set environment variables
ENV NODE_ENV production

# Copy package.json and package-lock.json files to the container
COPY package*.json ./

# Install Node.js dependencies
RUN npm install

# Install PM2 globally
RUN npm install -g pm2

# Copy all application files to the container
COPY . .

# Update package lists and install Python and required system packages
RUN apt-get update && apt-get install -y \
    python3 \
    python3-pip \
    && rm -rf /var/lib/apt/lists/*

# Install Python packages from requirements.txt if available
COPY requirements.txt ./
RUN pip3 install --no-cache-dir -r requirements.txt

# Expose the application port
EXPOSE 3001

# Define the command to run your application
CMD ["pm2-runtime", "server.js"]
