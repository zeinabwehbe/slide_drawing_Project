# Teacher Dashboard - View Previous Sessions

## Overview
Teachers can now view all their previous sessions and student responses directly from the PowerPoint add-in. This feature ensures that **each teacher only sees their own sessions** - data is fully isolated per teacher account.

## Features

### 1. **View Session History**
- Click "📚 View Previous Sessions" button on the main screen
- See all your past activities with:
  - Session date and time
  - Status (Active/Closed)
  - Timer settings
  - Number of submissions
  - Privacy settings (hide names, allow resubmit)

### 2. **View Session Details**
- Click on any session card to see full details
- View all student submissions with:
  - Student names
  - Drawing images
  - Submission timestamps
  - Session configuration

### 3. **Security & Privacy**
- **Teacher-specific data**: You only see YOUR sessions
- **Database-level filtering**: Sessions filtered by teacher_id
- **JWT authentication**: All API calls require valid token
- **Automatic logout**: Sessions expire after 7 days

## How to Use

### Step 1: Login
1. Open the add-in in PowerPoint
2. Login with your teacher credentials
3. You'll see the main activity creation screen

### Step 2: Access History
1. Click "📚 View Previous Sessions" button
2. Browse your session history (sorted by most recent first)
3. Each card shows:
   - 📅 Date/Time created
   - 🆔 Session ID
   - ✅/❌ Active or Closed status
   - ⏱️ Timer setting
   - 📝 Number of submissions
   - 👁️ Hide names setting

### Step 3: View Submissions
1. Click any session card
2. See session details (timer, settings, etc.)
3. Scroll through all student submissions
4. Each submission shows:
   - Student name
   - Drawing image
   - Timestamp

### Step 4: Navigate Back
- Click "← Back" to return to session list
- Click "← Back" again to return to main screen

## API Endpoints

### GET `/api/teacher/sessions`
Fetch all sessions for authenticated teacher.

**Authentication**: Required (JWT Bearer token)

**Response:**
```json
{
  "success": true,
  "sessions": [
    {
      "id": "abc123",
      "timer_duration": 5,
      "hide_names": false,
      "allow_resubmit": true,
      "start_time": 1700000000000,
      "timer_end": 1700000300000,
      "closed": true,
      "created_at": "2024-11-22T10:30:00.000Z",
      "submission_count": 15
    }
  ]
}
```

### GET `/api/teacher/sessions/:sessionId/submissions`
Fetch session details and all submissions (teacher-owned sessions only).

**Authentication**: Required (JWT Bearer token)

**Access Control**: Returns 403 if session doesn't belong to authenticated teacher

**Response:**
```json
{
  "success": true,
  "session": {
    "id": "abc123",
    "slideImage": "data:image/png;base64,...",
    "timerDuration": 5,
    "hideNames": false,
    "allowResubmit": true,
    "startTime": 1700000000000,
    "timerEnd": 1700000300000,
    "closed": true,
    "createdAt": "2024-11-22T10:30:00.000Z"
  },
  "submissions": [
    {
      "id": "sub1",
      "studentName": "John Doe",
      "drawingData": "data:image/png;base64,...",
      "timestamp": "2024-11-22T10:35:00.000Z"
    }
  ]
}
```

## Database Schema

### Sessions Table
```sql
CREATE TABLE sessions (
  id VARCHAR(20) PRIMARY KEY,
  teacher_id INT NOT NULL,  -- Links to teachers table
  slide_image LONGTEXT NOT NULL,
  -- ... other fields
  FOREIGN KEY (teacher_id) REFERENCES teachers(id) ON DELETE CASCADE
);
```

### Security Implementation
```javascript
// Server-side verification
const teacherId = req.user.id;  // From JWT token

// Only fetch sessions for this teacher
const [sessions] = await db.query(
  'SELECT * FROM sessions WHERE teacher_id = ?',
  [teacherId]
);
```

## UI Components

### History Section (`historySection`)
- List of session cards
- Sorted by created date (newest first)
- Color-coded status indicators
- Click to view details

### Session Detail Section (`sessionDetailSection`)
- Session metadata
- Submission count
- List of all student drawings
- Scrollable if many submissions

### Navigation
- "← Back" buttons for easy navigation
- Maintains authentication state
- Smooth section transitions

## JavaScript Functions

### `showHistory()`
Fetches and displays teacher's session list.

### `displaySessionsList(sessions)`
Renders session cards with data.

### `viewSessionDetails(sessionId)`
Fetches and displays specific session + submissions.

### `displaySessionDetails(session, submissions)`
Renders session info and submission gallery.

## Testing

### 1. Create Test Data
```powershell
# Seed teacher account
Get-Content seed.sql | mysql -u root -pYourPassword
```

### 2. Create Some Sessions
- Login as teacher
- Create 2-3 activities with different settings
- Have students submit drawings

### 3. Test History View
- Click "View Previous Sessions"
- Verify all your sessions appear
- Click on a session
- Verify submissions display correctly

### 4. Test Security
- Logout and login as different teacher
- Create a new session
- Verify you DON'T see the other teacher's sessions

## Troubleshooting

### Sessions not loading
- Check authentication (valid JWT token)
- Verify database connection
- Check browser console for errors
- Ensure `authToken` is set in localStorage

### "Access denied" error
- Trying to access another teacher's session
- Re-login to get fresh token
- Check teacher_id in sessions table

### Submissions not showing
- Verify session has submissions in database
- Check `drawings` table for session_id
- Verify foreign key relationships

### Images not displaying
- Check `drawing_data` format (should be base64)
- Verify image data not truncated
- Check browser console for image errors

## Future Enhancements
- Export submissions as PDF
- Filter/search sessions by date
- Sort submissions by name/time
- Download individual drawings
- Bulk operations (delete old sessions)
- Analytics dashboard (avg submissions, completion rates)
- Share sessions with other teachers
