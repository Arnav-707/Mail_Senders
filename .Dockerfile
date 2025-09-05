FROM node:18
WORKDIR /app

COPY package*.json ./
RUN npm install

# Copy all your project files into the container
# This includes index.js, .env, contacts.xlsx, etc.
COPY . .

# EXPOSE the port your Express server is listening on
EXPOSE 3001

# The command to run your main script file
CMD ["node", "app.js"]

