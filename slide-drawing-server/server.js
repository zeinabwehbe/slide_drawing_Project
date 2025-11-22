const express = require('express');
const http = require('http');
const socketIo = require('socket.io');
const path = require('path');
const cors = require('cors');
const bcrypt = require('bcryptjs');
const jwt = require('jsonwebtoken');
const db = require('./db');

const JWT_SECRET = process.env.JWT_SECRET || '***REMOVED***';

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

// Authentication middleware
const authenticateToken = (req, res, next) => {
  const authHeader = req.headers['authorization'];
  const token = authHeader && authHeader.split(' ')[1];

  if (!token) {
    return res.status(401).json({ error: 'Access token required' });
  }

  jwt.verify(token, JWT_SECRET, (err, user) => {
    if (err) {
      return res.status(403).json({ error: 'Invalid or expired token' });
    }
    req.user = user;
    next();
  });
};

// Register endpoint
app.post('/api/auth/register', async (req, res) => {
  try {
    const { username, email, password } = req.body;

    if (!username || !email || !password) {
      return res.status(400).json({ error: 'Username, email, and password are required' });
    }

    // Check if user already exists
    const [existingUsers] = await db.query(
      'SELECT id FROM teachers WHERE username = ? OR email = ?',
      [username, email]
    );

    if (existingUsers.length > 0) {
      return res.status(409).json({ error: 'Username or email already exists' });
    }

    // Hash password
    const passwordHash = await bcrypt.hash(password, 10);

    // Insert new teacher
    const [result] = await db.query(
      'INSERT INTO teachers (username, email, password_hash) VALUES (?, ?, ?)',
      [username, email, passwordHash]
    );

    // Generate token
    const token = jwt.sign(
      { id: result.insertId, username, email },
      JWT_SECRET,
      { expiresIn: '7d' }
    );

    res.json({
      success: true,
      token,
      user: { id: result.insertId, username, email }
    });
  } catch (error) {
    console.error('Registration error:', error);
    res.status(500).json({ error: 'Registration failed' });
  }
});

// Login endpoint
app.post('/api/auth/login', async (req, res) => {
  try {
    const { username, password } = req.body;

    if (!username || !password) {
      return res.status(400).json({ error: 'Username and password are required' });
    }

    // Find user
    const [users] = await db.query(
      'SELECT * FROM teachers WHERE username = ?',
      [username]
    );

    if (users.length === 0) {
      return res.status(401).json({ error: 'Invalid credentials' });
    }

    const user = users[0];

    // Verify password
    const validPassword = await bcrypt.compare(password, user.password_hash);

    if (!validPassword) {
      return res.status(401).json({ error: 'Invalid credentials' });
    }

    // Generate token
    const token = jwt.sign(
      { id: user.id, username: user.username, email: user.email },
      JWT_SECRET,
      { expiresIn: '7d' }
    );

    res.json({
      success: true,
      token,
      user: { id: user.id, username: user.username, email: user.email }
    });
  } catch (error) {
    console.error('Login error:', error);
    res.status(500).json({ error: 'Login failed' });
  }
});

// Verify token endpoint
app.get('/api/auth/verify', authenticateToken, (req, res) => {
  res.json({ success: true, user: req.user });
});

// Get all sessions for a teacher
app.get('/api/teacher/sessions', authenticateToken, async (req, res) => {
  try {
    const teacherId = req.user.id;
    
    const [sessions] = await db.query(
      `SELECT 
        s.id, 
        s.timer_duration, 
        s.hide_names, 
        s.allow_resubmit, 
        s.start_time, 
        s.timer_end, 
        s.closed, 
        s.created_at,
        COUNT(DISTINCT d.id) as submission_count
      FROM sessions s
      LEFT JOIN drawings d ON s.id = d.session_id
      WHERE s.teacher_id = ?
      GROUP BY s.id
      ORDER BY s.created_at DESC`,
      [teacherId]
    );
    
    res.json({ success: true, sessions });
  } catch (error) {
    console.error('Error fetching teacher sessions:', error);
    res.status(500).json({ error: 'Failed to fetch sessions' });
  }
});

// Get all submissions for a specific session (teacher only)
app.get('/api/teacher/sessions/:sessionId/submissions', authenticateToken, async (req, res) => {
  try {
    const teacherId = req.user.id;
    const sessionId = req.params.sessionId;
    
    // Verify session belongs to this teacher
    const [sessions] = await db.query(
      'SELECT * FROM sessions WHERE id = ? AND teacher_id = ?',
      [sessionId, teacherId]
    );
    
    if (sessions.length === 0) {
      return res.status(403).json({ error: 'Access denied or session not found' });
    }
    
    const session = sessions[0];
    
    // Get all drawings for this session
    const [drawings] = await db.query(
      'SELECT id, student_name, drawing_data, timestamp FROM drawings WHERE session_id = ? ORDER BY timestamp DESC',
      [sessionId]
    );
    
    res.json({
      success: true,
      session: {
        id: session.id,
        slideImage: session.slide_image,
        timerDuration: session.timer_duration,
        hideNames: session.hide_names,
        allowResubmit: session.allow_resubmit,
        startTime: session.start_time,
        timerEnd: session.timer_end,
        closed: session.closed,
        createdAt: session.created_at
      },
      submissions: drawings.map(d => ({
        id: d.id,
        studentName: d.student_name,
        drawingData: d.drawing_data,
        timestamp: d.timestamp
      }))
    });
  } catch (error) {
    console.error('Error fetching session submissions:', error);
    res.status(500).json({ error: 'Failed to fetch submissions' });
  }
});

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
      const { slideImage, settings, teacherId } = data;
      
      if (!teacherId) {
        socket.emit('session-error', { message: 'Teacher authentication required' });
        return;
      }
      
      const timerDuration = settings.timerDuration || null;
      const startTime = Date.now();
      const timerEnd = timerDuration ? startTime + (timerDuration * 60 * 1000) : null;
      
      // Insert session into database
      await db.query(
        `INSERT INTO sessions 
        (id, teacher_id, slide_image, teacher_socket_id, timer_duration, hide_names, allow_resubmit, start_time, timer_end, closed) 
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
        [
          sessionId,
          teacherId,
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