# 📸 Mostafa Khirallah Photography Client Delivery Portal

A secure, premium-grade, and cinematic web application for photographer **Mostafa Khirallah** to deliver photo sessions directly to clients. The site acts as a client dashboard and favorite photo selection portal.

## 🛠️ Architecture & Technology Stack

- **Frontend**: Single-Page Application (SPA) hosted on **GitHub Pages**, styled with **Tailwind CSS v3** (Obsidian dark theme, gold accents, glassmorphic panels, and smooth micro-animations).
- **Backend API**: **Google Apps Script** deployed as a CORS-compliant Web App API.
- **Database**: **Google Sheets** (Client data matching, favorite logs, and notes).
- **Asset Storage**: **Google Drive** (Session photos retrieved dynamically via high-res Google CDN thumbnail URLs).

```
[ Client Browser / GitHub Pages ]
       │
       ▼ (HTTP GET Fetch)
[ Google Apps Script Web App API ]
       │
       ├─► [ Google Sheets (Database) ]
       └─► [ Google Drive (Photos Storage) ]
```

## ✨ Features

- **Secure Client Access**: Login page validating Mobile Number and Invoice Number against a Google Sheet database.
- **Cinematic Dynamic Gallery**: Masonry-like image grid loading photos dynamically with responsive skeleton indicators.
- **Intellectual Property Protection**: Blocked context menus (right-click) and image dragging on proofing photos.
- **Favorites Selection Drawer**: Interactive "Heart" selection allowing clients to select their favorite proof photos, add adjustment comments/notes, and submit them back to the photographer's Sheet.
- **Direct Downloads**: Direct high-res download buttons per photo.
- **Passcode-Protected Admin Panel**: Direct client session management (create, update, delete) within the web application.

---

## 🚀 Setup & Linking Guide

For complete instructions on how to link your Google Apps Script API with this frontend and launch your GitHub Pages host, please refer to the detailed guide:
👉 **[GitHub Pages Deployment Guide (github_deployment.md)](./github_deployment.md)**
