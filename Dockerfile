# 1. ใช้ Node.js เวอร์ชันเล็กๆ (Alpine) เพื่อประหยัดพื้นที่
FROM node:18-alpine

# 2. กำหนดโฟลเดอร์ทำงานใน Docker
WORKDIR /app

# 3. Copy ไฟล์ package.json ไปก่อน (เพื่อ Cache layer)
COPY package*.json ./

# 4. ติดตั้ง dependencies
RUN npm install --production

# 5. Copy โค้ดทั้งหมดเข้าไป
COPY . .

# 6. บอกว่าจะใช้ Port 3000
EXPOSE 3000

# 7. คำสั่งรันเมื่อเริ่ม Container
CMD ["npm", "start"]