FROM node:20-alpine

WORKDIR /app

COPY package*.json ./
RUN npm ci

COPY . .
RUN npm run build

RUN mkdir -p /data

EXPOSE 4000

CMD ["npm", "run", "serve"]
