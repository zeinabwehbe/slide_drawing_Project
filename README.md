# Slide Drawing Project - MySQL Setup Guide

## Prerequisites
- Node.js (v14 or higher)
- MySQL Server (v5.7 or higher)

## Setup Instructions

### 1. Install MySQL (if not already installed)
Download and install MySQL from: https://dev.mysql.com/downloads/mysql/

### 2. Create the Database
Open MySQL command line or MySQL Workbench and run:

```sql
mysql -u root -p
```

Then execute the schema file:
```sql
source c:/Users/Zayna/Downloads/RM_Project/slide-drawing-server/schema.sql
```

Or manually run the SQL commands from `schema.sql` file.

### 3. Configure Database Connection
Copy `.env.example` to `.env` in the `slide-drawing-server` folder, then fill in your values:

```
DB_HOST=localhost
DB_USER=root
DB_PASSWORD=your_mysql_password_here
DB_NAME=slide_drawing_db
PORT=3001
JWT_SECRET=replace_with_a_long_random_value
```

**Important:** Replace `your_mysql_password` with your actual MySQL root password.

### 4. Install Dependencies
Dependencies are already installed. If needed, run:
```bash
cd slide-drawing-server
npm install
```

### 5. Start the Server
```bash
cd slide-drawing-server
npm start
```

The server will start on http://localhost:3001

### 6. Run the PowerPoint Add-in
In a separate terminal:
```bash
cd SlideDrawingProject
npm install
npm start
```

This will start the Office Add-in for PowerPoint.

## Project Structure

### Backend (slide-drawing-server)
- `server.js` - Main server file with Socket.IO and MySQL integration
- `db.js` - Database connection configuration
- `schema.sql` - Database schema
- `.env` - Environment variables for database credentials
- `public/` - Static HTML files for student and viewer interfaces

### Frontend (SlideDrawingProject)
- PowerPoint Add-in built with Office.js
- Connects to the backend server for real-time drawing sessions

## Database Tables

### sessions
- Stores session information (slide images, settings, timers)

### drawings
- Stores student submissions linked to sessions

## Features
- Real-time drawing submissions
- Timer support with auto-end
- Resubmission controls
- Anonymous mode (hide names)
- All data persisted in MySQL database

## Troubleshooting

### Cannot connect to MySQL
- Ensure MySQL service is running
- Check credentials in `.env` file
- Verify database `slide_drawing_db` exists

### Port conflicts
- Change PORT in `.env` if 3001 is in use
- Update the server URL in PowerPoint add-in if needed

### Dependencies issues
```bash
npm install
```

## Testing the Connection
The server will log connection status on startup:
- ✅ MySQL Database connected successfully - Good!
- ❌ Error connecting to MySQL database - Check your configuration


(for backend and login)
```
npm start
``` 