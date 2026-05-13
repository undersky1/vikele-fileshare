# Vikele FileShare

Vikele FileShare is a self-hosted file sharing platform, built and maintained by **Vikele Solutions**. Based on the open-source Pingvin Share project (BSD-2-Clause).

## Features

- Share files using a link
- Unlimited file size (restricted only by disk space)
- Set an expiration date for shares
- Secure shares with visitor limits and passwords
- Email recipients
- Reverse shares
- OIDC and LDAP authentication
- Integration with ClamAV for security scans
- Multiple storage providers: **Local storage**, **S3**, and **SharePoint**

## Setup

### Installation with Docker (recommended)

1. Clone this repository
2. Run `docker compose -f docker-compose.production.yml up -d --build`

The website is now listening on `http://localhost:3000`.

### SharePoint Storage

Vikele FileShare supports storing uploaded files directly in a SharePoint document library. To configure:

1. Register an Azure AD (Entra ID) app with `Sites.ReadWrite.All` application permission
2. Navigate to **Admin → SharePoint** in the web UI
3. Enter your Tenant ID, Client ID, Client Secret, Site ID, and Drive ID
4. Enable the SharePoint provider

See the [SharePoint Setup Guide](docs/docs/setup/sharepoint.md) for detailed instructions.

## Documentation

For more installation options and advanced configurations, please refer to the documentation.

## Setup project (development)

### Backend

1. Open the `backend` folder
2. Install the dependencies with `npm install`
3. Push the database schema to the database by running `npx prisma db push`
4. Seed the database with `npx prisma db seed`
5. Start the backend with `npm run dev`

### Frontend

1. Start the backend first
2. Open the `frontend` folder
3. Install the dependencies with `npm install`
4. Start the frontend with `npm run dev`

---

© Vikele Solutions — [vikele.co.za](https://vikele.co.za)
