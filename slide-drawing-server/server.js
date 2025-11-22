const express = require('express');
const http = require('http');
const socketIo = require('socket.io');
const path = require('path');
const cors = require('cors');
const db = require('./db');

const app = express();
app.use(cors());
const server = http.createServer(app);
const io = socketIo(server, {
  cors: {
    origin: "*",
    methods: ["GET", "POST"]
  }
});

app.use(express.static('public'));
app.use(express.json({limit: '10mb'})); // Reduced from 50mb

// Serve student interface
app.get('/student/:sessionId', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'student.html'));
});

// Serve viewer interface
app.get('/viewer', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'viewer.html'));
});

// API to get session data (for student page)
app.get('/api/session/:sessionId', async (req, res) => {
  try {
    const [sessions] = await db.query(
      'SELECT * FROM sessions WHERE id = ?',
      [req.params.sessionId]
    );
    
    if (sessions.length > 0) {
      const session = sessions[0];
      const [drawings] = await db.query(
        'SELECT * FROM drawings WHERE session_id = ?',
        [req.params.sessionId]
      );
      
      res.json({
        slideImage: session.slide_image,
        drawings: drawings.map(d => ({
          id: d.id,
          studentName: d.student_name,
          drawingData: d.drawing_data,
          timestamp: d.timestamp
        })),
        settings: {
          timerDuration: session.timer_duration,
          hideNames: session.hide_names,
          allowResubmit: session.allow_resubmit,
          startTime: session.start_time
        },
        startTime: session.start_time,
        timerEnd: session.timer_end,
        closed: session.closed
      });
    } else {
      res.status(404).json({ error: 'Session not found' });
    }
  } catch (error) {
    console.error('Error fetching session:', error);
    res.status(500).json({ error: 'Database error' });
  }
});

// Socket.io for real-time communication
io.on('connection', (socket) => {
  console.log('Client connected:', socket.id);
  
  // Teacher creates session
  socket.on('create-session', async (data) => {
    try {
      const sessionId = Math.random().toString(36).substring(7);
      const { slideImage, settings } = data;
      
      const timerDuration = settings.timerDuration || null;
      const startTime = Date.now();
      const timerEnd = timerDuration ? startTime + (timerDuration * 60 * 1000) : null;
      
      // Insert session into database
      await db.query(
        `INSERT INTO sessions 
        (id, slide_image, teacher_socket_id, timer_duration, hide_names, allow_resubmit, start_time, timer_end, closed) 
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`,
        [
          sessionId,
          slideImage,
          socket.id,
          timerDuration,
          settings.hideNames || false,
          settings.allowResubmit !== undefined ? settings.allowResubmit : true,
          startTime,
          timerEnd,
          false
        ]
      );
      
      console.log('Session created:', sessionId, {
        timerDuration: timerDuration,
        hideNames: settings.hideNames,
        allowResubmit: settings.allowResubmit
      });
      
      // Join teacher to session room
      socket.join(`teacher-${sessionId}`);
      
      // Send response back to teacher WITH TIMER INFO
      socket.emit('session-created', { 
        sessionId,
        timerDuration: timerDuration,
        startTime: startTime
      });
      
      // Auto-end session when timer expires
      if (timerDuration) {
        setTimeout(async () => {
          try {
            const [sessions] = await db.query(
              'SELECT closed FROM sessions WHERE id = ?',
              [sessionId]
            );
            
            if (sessions.length > 0 && !sessions[0].closed) {
              await db.query(
                'UPDATE sessions SET closed = ? WHERE id = ?',
                [true, sessionId]
              );
              
              // Notify everyone
              io.to(`teacher-${sessionId}`).emit('session-ended');
              io.to(`student-${sessionId}`).emit('session-ended');
              io.emit('session-ended-viewer', { sessionId });
              
              console.log('Session auto-ended (timer expired):', sessionId);
            }
          } catch (error) {
            console.error('Error auto-ending session:', error);
          }
        }, timerDuration * 60 * 1000);
      }
      
    } catch (error) {
      console.error('Error creating session:', error);
      socket.emit('session-error', { message: error.message });
    }
  });
  
  // Teacher joins session (for reconnection)
  socket.on('teacher-join', (sessionId) => {
    socket.join(`teacher-${sessionId}`);
    console.log('Teacher joined session:', sessionId);
  });
  
  // Student joins session
  socket.on('student-join', (sessionId) => {
    socket.join(`student-${sessionId}`);
    console.log('Student joined session:', sessionId);
  });
  
  // Student submits drawing
  socket.on('submit-drawing', async (data) => {
    try {
      const { sessionId, drawingData, studentName } = data;
      
      // Get session from database
      const [sessions] = await db.query(
        'SELECT * FROM sessions WHERE id = ?',
        [sessionId]
      );
      
      if (sessions.length === 0) {
        socket.emit('submission-confirmed', { 
          success: false, 
          error: 'Session not found or has ended' 
        });
        return;
      }
      
      const session = sessions[0];
      
      // Check if session is closed
      if (session.closed) {
        socket.emit('submission-confirmed', { 
          success: false, 
          error: 'Activity has ended. Submissions are no longer accepted.' 
        });
        return;
      }
      
      // Check timer
      if (session.timer_end && Date.now() > session.timer_end) {
        socket.emit('submission-confirmed', { 
          success: false, 
          error: 'Time has expired for this activity.' 
        });
        return;
      }
      
      // Check if student already submitted
      const [existingSubmissions] = await db.query(
        'SELECT * FROM drawings WHERE session_id = ? AND student_name = ?',
        [sessionId, studentName]
      );
      
      if (existingSubmissions.length > 0) {
        // Check if resubmit is allowed
        if (!session.allow_resubmit) {
          socket.emit('submission-confirmed', { 
            success: false, 
            error: 'Resubmissions are not allowed for this activity.' 
          });
          return;
        }
        
        // Update existing submission
        const existingSubmission = existingSubmissions[0];
        await db.query(
          'UPDATE drawings SET drawing_data = ?, timestamp = NOW() WHERE id = ?',
          [drawingData, existingSubmission.id]
        );
        
        console.log('Drawing updated for session:', sessionId, 'by', studentName);
        
        // Notify teacher of update
        const updatedSubmission = {
          id: existingSubmission.id,
          studentName: studentName,
          drawingData: drawingData,
          timestamp: new Date()
        };
        io.to(`teacher-${sessionId}`).emit('submission-updated', updatedSubmission);
        
        socket.emit('submission-confirmed', { 
          success: true, 
          message: 'Submission updated successfully!' 
        });
      } else {
        // Create new submission
        const submissionId = Math.random().toString(36).substring(7);
        await db.query(
          'INSERT INTO drawings (id, session_id, student_name, drawing_data) VALUES (?, ?, ?, ?)',
          [submissionId, sessionId, studentName || 'Anonymous', drawingData]
        );
        
        console.log('Drawing submitted for session:', sessionId, 'by', studentName || 'Anonymous');
        
        // Notify teacher
        const submission = {
          id: submissionId,
          studentName: studentName || 'Anonymous',
          drawingData: drawingData,
          timestamp: new Date()
        };
        io.to(`teacher-${sessionId}`).emit('new-submission', submission);
        
        socket.emit('submission-confirmed', { 
          success: true,
          message: 'Drawing submitted successfully!'
        });
      }
    } catch (error) {
      console.error('Error submitting drawing:', error);
      socket.emit('submission-confirmed', { 
        success: false, 
        error: 'Database error occurred' 
      });
    }
  });
  
  // Get all submissions
  socket.on('get-submissions', async (sessionId) => {
    try {
      const [drawings] = await db.query(
        'SELECT * FROM drawings WHERE session_id = ?',
        [sessionId]
      );
      
      const submissions = drawings.map(d => ({
        id: d.id,
        studentName: d.student_name,
        drawingData: d.drawing_data,
        timestamp: d.timestamp
      }));
      
      socket.emit('all-submissions', submissions);
    } catch (error) {
      console.error('Error getting submissions:', error);
      socket.emit('all-submissions', []);
    }
  });
  
  // End session
  socket.on('end-session', async (sessionId) => {
    try {
      const [sessions] = await db.query(
        'SELECT * FROM sessions WHERE id = ?',
        [sessionId]
      );
      
      if (sessions.length > 0) {
        // Mark session as closed
        await db.query(
          'UPDATE sessions SET closed = ? WHERE id = ?',
          [true, sessionId]
        );
        
        // Notify all students in this session
        io.to(`student-${sessionId}`).emit('session-ended');
        
        console.log('Session ended:', sessionId);
        
        // Optional: Delete session after some time (kept in database for records)
        // You can add cleanup logic here if needed
      }
    } catch (error) {
      console.error('Error ending session:', error);
    }
  });
  
  socket.on('disconnect', () => {
    console.log('Client disconnected:', socket.id);
  });
});

const PORT = process.env.PORT || 3001;
server.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
  console.log('Socket.io enabled for real-time communication');
});