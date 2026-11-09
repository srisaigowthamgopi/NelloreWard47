# Deployment Guide - Ward 47 Municipal Citizen Portal & Admin Dashboard

This guide explains how to deploy the **47th Municipal Corporation Ward Citizen Portal & Admin Dashboard** across different environments, from local office networks to free cloud hosting and production cloud servers.

---

## 📋 Project Overview & Architecture

- **Backend**: Node.js + Express (`server.js`)
- **Frontend**: HTML5 + CSS3 + Vanilla JavaScript (`index.html`, `admin.html`, `style.css`)
- **Storage**: Excel Spreadsheet (`data/Ward47_Public_Problems.xlsx` via `exceljs`) + Uploaded images (`uploads/`)
- **Key Features**:
  - One-click Telugu & English language switcher with `localStorage` persistence.
  - Public Citizen Grievance Submission with photo attachments.
  - Dedicated Admin Dashboard with secret key authentication, live status updates, search/filter, and Excel export.

---

## ⚙️ Prerequisites

1. **Install Node.js** (LTS version 18.x or 20.x):
   - Download and install from [nodejs.org](https://nodejs.org/).
   - Verify installation in terminal:
     ```bash
     node -v
     npm -v
     ```

2. **Install Dependencies**:
   Open a terminal in the project folder and run:
   ```bash
   npm install
   ```

---

## 🚀 Option 1: Cloud Deployment on Render.com (Recommended)

Render provides Node.js hosting with automatic SSL (HTTPS). This application stores complaints and photos on disk, so this configuration uses a paid web service with a persistent disk; the free service tier does not provide persistent application storage.

### Steps:
1. **Push to GitHub**:
   - Create a private GitHub repository for the project. Do not commit `.env`, complaint data, or uploaded photos.
   - In a terminal in the project folder, initialize Git if needed and push the source:
     ```bash
     git init
     git add .
     git commit -m "Prepare Ward 47 portal for deployment"
     git branch -M main
     git remote add origin https://github.com/<your-username>/ward47-portal.git
     git push -u origin main
     ```

2. **Create the Render service from the Blueprint**:
   - Sign up or log in at [Render.com](https://render.com).
   - Choose **New + ➔ Blueprint** and connect the GitHub repository.
   - Render reads `render.yaml`, which configures the Node service, health check, Singapore region, and a 1 GB persistent disk mounted at `/var/data`. Review Render's current pricing before creating resources.

3. **Configure Environment Variables**:
   - Set `ADMIN_KEY` in the Render dashboard to a unique, long random initial secret. Do not use the example value or commit the secret to GitHub.
   - Render supplies `PORT`; the Blueprint sets `NODE_ENV`, `DATA_DIR`, and `UPLOAD_DIR`.
   - Password changes made through the admin dashboard are stored on the persistent disk and take precedence over `ADMIN_KEY`.

4. **Deploy and verify**:
   - Create the Blueprint resources and wait for the service to become healthy.
   - Open `https://<service-name>.onrender.com` for the citizen portal and `/admin` for the admin dashboard.
   - Confirm `/api/health` reports `{"status":"ok"}`. Confirm `/data/Ward47_Public_Problems.xlsx` returns 404; complaint records are only available through authenticated admin endpoints.

The persistent disk keeps the Excel workbook, admin key file, and uploaded images across service restarts. Back up the workbook and photos regularly; a disk is not a backup.

---

## 🌐 Option 2: Deploy on Railway.app

1. Go to [railway.app](https://railway.app) and sign up with GitHub.
2. Click **"New Project"** ➔ **"Deploy from GitHub repo"**.
3. Select your repository.
4. Add Environment Variables:
   - `ADMIN_KEY` = `your_strong_admin_key`
5. Railway will automatically detect `package.json`, run `npm install`, and start `node server.js`.
6. Go to **Settings ➔ Generate Domain** to get a public `https://...up.railway.app` URL.

---

## 🏢 Option 3: Local Municipal / Sachivalayam Office Deployment (Windows or Linux PC)

If this website needs to run within the local Ward Office or Sachivalayam for staff and local visitors over LAN or WiFi:

### 1. Run the Server
Open PowerShell or Command Prompt in the folder:
```powershell
npm install
node server.js
```

### 2. Run 24/7 in Background using PM2
To keep the server running even if the terminal window closes or after a computer reboot:
```bash
npm install -g pm2
pm2 start server.js --name "ward47-portal"
pm2 save
pm2 startup
```

### 3. Access Over Office Local Network (LAN / Wi-Fi)
- Find your local IP address:
  - Windows: Run `ipconfig` (e.g. `192.168.1.50`)
  - Linux: Run `hostname -I`
- Any computer or mobile phone connected to the same office Wi-Fi can open:
  - Citizen Portal: `http://192.168.1.50:3000`
  - Admin Dashboard: `http://192.168.1.50:3000/admin`

---

## ☁️ Option 4: Production Cloud VPS (DigitalOcean, AWS EC2, Hostinger, Linode)

For a custom domain (e.g. `ward47nellore.in`) and high-traffic production use:

### 1. Server Setup (Ubuntu Linux)
```bash
# Update server packages
sudo apt update && sudo apt upgrade -y

# Install Node.js
curl -fsSL https://deb.nodesource.com/setup_20.x | sudo -E bash -
sudo apt install -y nodejs nginx git

# Clone project
git clone https://github.com/<your-username>/ward47-portal.git /var/www/ward47
cd /var/www/ward47
npm install

# Start with PM2
sudo npm install -g pm2
ADMIN_KEY="yourSecureSecretKey" pm2 start server.js --name "ward47-portal"
pm2 save
sudo env PATH=$PATH:/usr/bin pm2 startup systemd -u $USER --hp $HOME
```

### 2. Configure Nginx Reverse Proxy
Edit `/etc/nginx/sites-available/ward47`:
```nginx
server {
    server_name yourdomain.com www.yourdomain.com;

    client_max_body_size 10M;

    location / {
        proxy_pass http://127.0.0.1:3000;
        proxy_http_version 1.1;
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection 'upgrade';
        proxy_set_header Host $host;
        proxy_cache_bypass $http_upgrade;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
    }
}
```

Enable site and restart Nginx:
```bash
sudo ln -s /etc/nginx/sites-available/ward47 /etc/nginx/sites-enabled/
sudo nginx -t
sudo systemctl restart nginx
```

### 3. Install Free SSL (HTTPS) with Certbot
```bash
sudo apt install -y certbot python3-certbot-nginx
sudo certbot --nginx -d yourdomain.com -d www.yourdomain.com
```

---

## 🔐 Admin Dashboard Security Checklist

1. **Custom Admin Password**:
   - In production, set a unique secret as the `ADMIN_KEY` environment variable in the hosting dashboard:
     ```bash
     ADMIN_KEY="<your-long-random-secret>"
     ```
   - The server refuses to start in production without this variable; do not rely on first-use password setup for a public service.
2. **Back Up Excel Data**:
   - The Excel spreadsheet is stored at `data/Ward47_Public_Problems.xlsx`.
   - Administrators can download a full backup anytime by clicking the **"📊 Export Excel"** button in the Admin Dashboard.
