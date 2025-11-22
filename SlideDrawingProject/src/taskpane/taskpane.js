/* global Office, io */

let socket;
let currentSessionId = null;
let uploadedSlideImage = null;
let timerInterval = null;
let currentUser = null;
let authToken = null;
const SERVER_URL = 'http://localhost:3001';

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    setupEventListeners();
    checkAuthentication();
    connectToServer();
  }
});

function setupEventListeners() {
  // Auth listeners
  document.getElementById('loginBtn').onclick = handleLogin;
  document.getElementById('registerBtn').onclick = handleRegister;
  document.getElementById('showRegisterLink').onclick = (e) => {
    e.preventDefault();
    document.getElementById('loginForm').style.display = 'none';
    document.getElementById('registerForm').style.display = 'block';
    document.getElementById('authMessage').style.display = 'none';
  };
  document.getElementById('showLoginLink').onclick = (e) => {
    e.preventDefault();
    document.getElementById('registerForm').style.display = 'none';
    document.getElementById('loginForm').style.display = 'block';
    document.getElementById('authMessage').style.display = 'none';
  };
  document.getElementById('logoutBtn').onclick = handleLogout;
  document.getElementById('viewHistoryBtn').onclick = showHistory;
  document.getElementById('backToCreateBtn').onclick = () => showSection('startSection');
  document.getElementById('backToHistoryBtn').onclick = showHistory;
  
  // Activity listeners
  document.getElementById('startButton').onclick = startActivity;
  document.getElementById('endButton').onclick = endActivity;
  document.getElementById('copyUrlBtn').onclick = copyUrl;
  document.getElementById('refreshBtn').onclick = refreshSubmissions;
  document.getElementById('viewAllBtn').onclick = openSubmissionsWindow;
  document.getElementById('slideImageUpload').onchange = handleImageUpload;
  document.getElementById('removeImageBtn').onclick = removeImage;
  document.getElementById('captureSlideBtn').onclick = handleCaptureSlide;
  
  // Timer option toggle
  var timerRadios = document.getElementsByName('timerOption');
  for (var i = 0; i < timerRadios.length; i++) {
    timerRadios[i].onchange = function() {
      var timerInput = document.getElementById('timerMinutes');
      timerInput.disabled = this.value === 'unlimited';
    };
  }
  
  // Enter key handlers for auth forms
  document.getElementById('loginPassword').onkeypress = (e) => {
    if (e.key === 'Enter') handleLogin();
  };
  document.getElementById('registerPassword').onkeypress = (e) => {
    if (e.key === 'Enter') handleRegister();
  };
}

async function handleCaptureSlide() {
  try {
    showMessage('Generating slide placeholder...', 'info');
    const slideImage = await captureCurrentSlide();
    
    if (slideImage) {
      uploadedSlideImage = slideImage;
      document.getElementById('previewImg').src = uploadedSlideImage;
      document.getElementById('imagePreview').style.display = 'block';
      showMessage('Placeholder created! Students will see slide reference.', 'success');
    } else {
      showMessage('Failed to create placeholder', 'error');
    }
  } catch (error) {
    console.error('Error capturing slide:', error);
    showMessage('Error: ' + error.message, 'error');
  }
}

function handleImageUpload(event) {
  const file = event.target.files[0];
  const MAX_SIZE = 10 * 1024 * 1024; // 10MB
  
  if (!file) return;
  
  if (!file.type.startsWith('image/')) {
    showMessage('Please upload an image file', 'error');
    return;
  }
  
  if (file.size > MAX_SIZE) {
    showMessage('Image too large. Maximum size is 10MB', 'error');
    return;
  }
  
  const reader = new FileReader();
  
  reader.onload = (e) => {
    uploadedSlideImage = e.target.result;
    document.getElementById('previewImg').src = uploadedSlideImage;
    document.getElementById('imagePreview').style.display = 'block';
  };
  
  reader.onerror = () => {
    showMessage('Failed to read image file', 'error');
  };
  
  reader.readAsDataURL(file);
}

function removeImage() {
  uploadedSlideImage = null;
  document.getElementById('slideImageUpload').value = '';
  const preview = document.getElementById('imagePreview');
  if (preview && preview.style) {
    preview.style.display = 'none';
  }
}

// Authentication Functions
function checkAuthentication() {
  const token = localStorage.getItem('authToken');
  const user = localStorage.getItem('authUser');
  
  if (token && user) {
    // Verify token with server
    fetch(`${SERVER_URL}/api/auth/verify`, {
      headers: {
        'Authorization': `Bearer ${token}`
      }
    })
    .then(response => response.json())
    .then(data => {
      if (data.success) {
        currentUser = JSON.parse(user);
        authToken = token;
        showSection('startSection');
        document.getElementById('loggedInUser').textContent = currentUser.username;
      } else {
        showSection('authSection');
      }
    })
    .catch(() => {
      showSection('authSection');
    });
  } else {
    showSection('authSection');
  }
}

async function handleLogin() {
  const username = document.getElementById('loginUsername').value.trim();
  const password = document.getElementById('loginPassword').value;
  
  if (!username || !password) {
    showAuthMessage('Please enter username and password', 'error');
    return;
  }
  
  try {
    const response = await fetch(`${SERVER_URL}/api/auth/login`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ username, password })
    });
    
    const data = await response.json();
    
    if (data.success) {
      currentUser = data.user;
      authToken = data.token;
      localStorage.setItem('authToken', data.token);
      localStorage.setItem('authUser', JSON.stringify(data.user));
      
      showSection('startSection');
      document.getElementById('loggedInUser').textContent = currentUser.username;
      showAuthMessage('Login successful!', 'success');
    } else {
      showAuthMessage(data.error || 'Login failed', 'error');
    }
  } catch (error) {
    showAuthMessage('Connection error. Make sure server is running.', 'error');
  }
}

async function handleRegister() {
  const username = document.getElementById('registerUsername').value.trim();
  const email = document.getElementById('registerEmail').value.trim();
  const password = document.getElementById('registerPassword').value;
  
  if (!username || !email || !password) {
    showAuthMessage('Please fill all fields', 'error');
    return;
  }
  
  if (password.length < 6) {
    showAuthMessage('Password must be at least 6 characters', 'error');
    return;
  }
  
  try {
    const response = await fetch(`${SERVER_URL}/api/auth/register`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ username, email, password })
    });
    
    const data = await response.json();
    
    if (data.success) {
      currentUser = data.user;
      authToken = data.token;
      localStorage.setItem('authToken', data.token);
      localStorage.setItem('authUser', JSON.stringify(data.user));
      
      showSection('startSection');
      document.getElementById('loggedInUser').textContent = currentUser.username;
      showAuthMessage('Registration successful!', 'success');
    } else {
      showAuthMessage(data.error || 'Registration failed', 'error');
    }
  } catch (error) {
    showAuthMessage('Connection error. Make sure server is running.', 'error');
  }
}

function handleLogout() {
  currentUser = null;
  authToken = null;
  localStorage.removeItem('authToken');
  localStorage.removeItem('authUser');
  showSection('authSection');
  
  // Clear form fields
  document.getElementById('loginUsername').value = '';
  document.getElementById('loginPassword').value = '';
}

function showAuthMessage(message, type) {
  const msgEl = document.getElementById('authMessage');
  msgEl.textContent = message;
  msgEl.style.display = 'block';
  msgEl.style.background = type === 'success' ? '#d4edda' : '#f8d7da';
  msgEl.style.color = type === 'success' ? '#155724' : '#721c24';
  msgEl.style.border = type === 'success' ? '1px solid #c3e6cb' : '1px solid #f5c6cb';
}

// Teacher History Functions
async function showHistory() {
  showSection('historySection');
  
  try {
    const response = await fetch(`${SERVER_URL}/api/teacher/sessions`, {
      headers: {
        'Authorization': `Bearer ${authToken}`
      }
    });
    
    const data = await response.json();
    
    if (data.success) {
      displaySessionsList(data.sessions);
    } else {
      document.getElementById('sessionsList').innerHTML = '<p style="color: #666; text-align: center;">Failed to load sessions</p>';
    }
  } catch (error) {
    document.getElementById('sessionsList').innerHTML = '<p style="color: #666; text-align: center;">Error loading sessions</p>';
  }
}

function displaySessionsList(sessions) {
  const listEl = document.getElementById('sessionsList');
  
  if (sessions.length === 0) {
    listEl.innerHTML = '<p style="color: #666; text-align: center;">No previous sessions found</p>';
    return;
  }
  
  listEl.innerHTML = sessions.map(session => {
    const date = new Date(session.created_at).toLocaleString();
    const status = session.closed ? '❌ Closed' : '✅ Active';
    const timer = session.timer_duration ? `${session.timer_duration} min` : 'No limit';
    
    return `
      <div class="session-card" data-session-id="${session.id}" style="background: white; border: 1px solid #e0e0e0; border-radius: 8px; padding: 15px; margin-bottom: 10px; cursor: pointer;">
        <div style="display: flex; justify-content: space-between; align-items: start; margin-bottom: 10px;">
          <div>
            <strong style="font-size: 14px;">${date}</strong>
            <br>
            <span style="font-size: 12px; color: #666;">ID: ${session.id}</span>
          </div>
          <span style="font-size: 12px; padding: 4px 8px; border-radius: 4px; background: ${session.closed ? '#fee' : '#efe'}; color: ${session.closed ? '#c33' : '#3c3'};">${status}</span>
        </div>
        <div style="font-size: 13px; color: #666;">
          <div style="margin: 5px 0;">⏱️ Timer: ${timer}</div>
          <div style="margin: 5px 0;">📝 Submissions: ${session.submission_count}</div>
          <div style="margin: 5px 0;">👁️ Hide Names: ${session.hide_names ? 'Yes' : 'No'}</div>
        </div>
      </div>
    `;
  }).join('');
  
  // Add click handlers to session cards
  document.querySelectorAll('.session-card').forEach(card => {
    card.addEventListener('click', function() {
      const sessionId = this.getAttribute('data-session-id');
      viewSessionDetails(sessionId);
    });
  });
}

async function viewSessionDetails(sessionId) {
  showSection('sessionDetailSection');
  document.getElementById('sessionInfo').innerHTML = '<p style="text-align: center; color: #666;">Loading session details...</p>';
  document.getElementById('detailSubmissionsList').innerHTML = '<p style="text-align: center; color: #666;">Loading submissions...</p>';
  
  try {
    const response = await fetch(`${SERVER_URL}/api/teacher/sessions/${sessionId}/submissions`, {
      headers: {
        'Authorization': `Bearer ${authToken}`
      }
    });
    
    const data = await response.json();
    
    if (response.ok && data.success) {
      displaySessionDetails(data.session, data.submissions);
    } else {
      console.error('Session details error:', data);
      document.getElementById('sessionInfo').innerHTML = `<p style="color: #c33;">${data.error || 'Failed to load session details'}</p>`;
      document.getElementById('detailSubmissionsList').innerHTML = '';
    }
  } catch (error) {
    console.error('Session details fetch error:', error);
    document.getElementById('sessionInfo').innerHTML = '<p style="color: #c33;">Error loading session details. Check console for details.</p>';
    document.getElementById('detailSubmissionsList').innerHTML = '';
  }
}

function displaySessionDetails(session, submissions) {
  const date = new Date(session.createdAt).toLocaleString();
  const status = session.closed ? '❌ Closed' : '✅ Active';
  const timer = session.timerDuration ? `${session.timerDuration} minutes` : 'No time limit';
  
  document.getElementById('sessionInfo').innerHTML = `
    <div style="margin-bottom: 10px;">
      <strong>Session ID:</strong> ${session.id}
    </div>
    <div style="margin-bottom: 10px;">
      <strong>Created:</strong> ${date}
    </div>
    <div style="margin-bottom: 10px;">
      <strong>Status:</strong> ${status}
    </div>
    <div style="margin-bottom: 10px;">
      <strong>Timer:</strong> ${timer}
    </div>
    <div style="margin-bottom: 10px;">
      <strong>Hide Names:</strong> ${session.hideNames ? 'Yes' : 'No'}
    </div>
    <div style="margin-bottom: 10px;">
      <strong>Allow Resubmit:</strong> ${session.allowResubmit ? 'Yes' : 'No'}
    </div>
  `;
  
  document.getElementById('detailSubmissionCount').textContent = submissions.length;
  
  const listEl = document.getElementById('detailSubmissionsList');
  
  if (submissions.length === 0) {
    listEl.innerHTML = '<p style="color: #666; text-align: center;">No submissions yet</p>';
    return;
  }
  
  listEl.innerHTML = submissions.map(sub => {
    const timestamp = new Date(sub.timestamp).toLocaleString();
    
    return `
      <div style="background: white; border: 1px solid #e0e0e0; border-radius: 8px; padding: 15px; margin-bottom: 10px;">
        <div style="display: flex; justify-content: space-between; align-items: start; margin-bottom: 10px;">
          <strong>${sub.studentName}</strong>
          <span style="font-size: 12px; color: #666;">${timestamp}</span>
        </div>
        <div style="text-align: center;">
          <img src="${sub.drawingData}" style="max-width: 100%; border-radius: 4px; border: 1px solid #e0e0e0;">
        </div>
      </div>
    `;
  }).join('');
}

function connectToServer() {
  socket = io(SERVER_URL);
  
  socket.on('connect', () => {
    console.log('Connected to server');
  });
  
  socket.on('session-created', (data) => {
    console.log('Session created:', data);
    currentSessionId = data.sessionId;
    
    // Update UI
    document.getElementById('sessionId').textContent = currentSessionId;
    document.getElementById('studentUrl').value = `${SERVER_URL}/student/${currentSessionId}`;
    document.getElementById('viewerUrl').value = `${SERVER_URL}/viewer?session=${currentSessionId}`;
    document.getElementById('copyViewerBtn').onclick = function() {
      const urlInput = document.getElementById('viewerUrl');
      urlInput.select();
      try {
        document.execCommand('copy');
        const btn = document.getElementById('copyViewerBtn');
        btn.textContent = 'Copied!';
        setTimeout(() => { btn.textContent = 'Copy'; }, 2000);
      } catch (err) {
        console.error('Copy failed:', err);
      }
    };
    
    // START TIMER DISPLAY - NEW!
    if (data.timerDuration) {
      startCountdownTimer(data.timerDuration, data.startTime);
    } else {
      startElapsedTimer(data.startTime);
    }
    
    showSection('activeSection');
    showMessage('Activity started successfully!', 'success');
    
    // Auto-open viewer if enabled
    if (window.activitySettings && window.activitySettings.autoOpenViewer) {
      setTimeout(() => openSubmissionsWindow(), 500);
    }
  });
  
  socket.on('session-error', (error) => {
    console.error('Session error:', error);
    showMessage('Failed to create session: ' + error.message, 'error');
    showSection('startSection');
  });
  
  socket.on('new-submission', (submission) => {
    addSubmissionToList(submission);
    updateSubmissionCount();
    showMessage('New submission received!', 'success');
  });
  
  socket.on('all-submissions', (submissions) => {
    displayAllSubmissions(submissions);
  });
  
  // Listen for session ended
  socket.on('session-ended', () => {
    if (timerInterval) {
      clearInterval(timerInterval);
      timerInterval = null;
    }
    showMessage('Activity has ended', 'error');
  });
  
  socket.on('connect_error', (error) => {
    console.error('Connection error:', error);
    showMessage('Cannot connect to server. Make sure it is running on port 3001.', 'error');
  });
}

// TIMER FUNCTIONS - NEW!
function startCountdownTimer(durationMinutes, startTime) {
  const endTime = startTime + (durationMinutes * 60 * 1000);
  const timerDisplay = document.getElementById('timerDisplay');
  
  timerInterval = setInterval(() => {
    const now = Date.now();
    const remaining = endTime - now;
    
    if (remaining <= 0) {
      clearInterval(timerInterval);
      timerInterval = null;
      timerDisplay.textContent = '⏱️ Time: 00:00 (Activity Ended)';
      timerDisplay.style.color = '#dc3545';
      return;
    }
    
    const minutes = Math.floor(remaining / 60000);
    const seconds = Math.floor((remaining % 60000) / 1000);
    timerDisplay.textContent = `⏱️ Time Remaining: ${minutes}:${seconds.toString().padStart(2, '0')}`;
    
    // Warning when < 1 minute
    if (remaining < 60000) {
      timerDisplay.style.color = '#dc3545';
    }
  }, 1000);
}

function startElapsedTimer(startTime) {
  const timerDisplay = document.getElementById('timerDisplay');
  
  timerInterval = setInterval(() => {
    const elapsed = Date.now() - startTime;
    const minutes = Math.floor(elapsed / 60000);
    const seconds = Math.floor((elapsed % 60000) / 1000);
    timerDisplay.textContent = `⏱️ Time Elapsed: ${minutes}:${seconds.toString().padStart(2, '0')}`;
  }, 1000);
}

async function startActivity() {
  console.log('Start Activity clicked');
  
  if (!socket || !socket.connected) {
    showMessage('Connecting to server...', 'error');
    connectToServer();
    setTimeout(() => {
      if (socket && socket.connected) {
        startActivity();
      } else {
        showMessage('Cannot connect to server. Make sure it is running on port 3001.', 'error');
      }
    }, 2000);
    return;
  }
  
  showSection('loadingSection');
  
  try {
    // Get settings
    const timerOption = document.querySelector('input[name="timerOption"]:checked').value;
    const timerDuration = timerOption === 'timed' ? parseInt(document.getElementById('timerMinutes').value) : null;
    const hideNames = document.getElementById('hideNames').checked;
    const autoOpenViewer = document.getElementById('autoOpenViewer').checked;
    const allowResubmit = document.getElementById('allowResubmit').checked;
    
    // Store settings globally
    window.activitySettings = {
      timerDuration: timerDuration,
      hideNames: hideNames,
      autoOpenViewer: autoOpenViewer,
      allowResubmit: allowResubmit
    };
    
    console.log('Settings:', window.activitySettings);
    
    // Get slide image
    console.log('Getting slide image...');
    const slideImage = uploadedSlideImage || await captureCurrentSlide();
    console.log('Slide image ready');
    
    // Create session with settings - INCLUDING TIMER AND TEACHER ID!
    console.log('Creating session via Socket.io...');
    socket.emit('create-session', { 
      slideImage: slideImage,
      teacherId: currentUser.id,
      settings: {
        timerDuration: timerDuration,
        hideNames: hideNames,
        allowResubmit: allowResubmit,
        startTime: Date.now()
      }
    });
    
  } catch (error) {
    console.error('Error starting activity:', error);
    showMessage('Failed to start activity: ' + error.message, 'error');
    showSection('startSection');
  }
}

async function captureCurrentSlide() {
  return new Promise(async (resolve, reject) => {
    try {
      // Try to get the current slide using PowerPoint API
      await PowerPoint.run(async (context) => {
        try {
          const slides = context.presentation.slides;
          slides.load("items");
          await context.sync();
          
          const selectedSlides = context.presentation.getSelectedSlides();
          selectedSlides.load("items");
          await context.sync();
          
          // Get the current slide index (1-based)
          let slideNumber = null;
          if (selectedSlides.items.length > 0) {
            // Find the index of the selected slide
            const selectedId = selectedSlides.items[0].id;
            const slideIndex = slides.items.findIndex(s => s.id === selectedId);
            slideNumber = slideIndex >= 0 ? slideIndex + 1 : null;
          }
          
          // For now, return styled placeholder with slide info
          // Note: PowerPoint API doesn't support direct image export yet
          resolve(createStyledPlaceholder(slideNumber));
        } catch (error) {
          console.log('PowerPoint API not available, using placeholder');
          resolve(createStyledPlaceholder());
        }
      });
    } catch (error) {
      console.log('Error accessing PowerPoint API:', error);
      // Fallback to placeholder
      resolve(createStyledPlaceholder());
    }
  });
}

function createStyledPlaceholder(slideNumber) {
  const canvas = document.createElement('canvas');
  canvas.width = 960;
  canvas.height = 540;
  const ctx = canvas.getContext('2d');
  
  const gradient = ctx.createLinearGradient(0, 0, 960, 540);
  gradient.addColorStop(0, '#667eea');
  gradient.addColorStop(1, '#764ba2');
  ctx.fillStyle = gradient;
  ctx.fillRect(0, 0, 960, 540);
  
  ctx.fillStyle = 'rgba(255, 255, 255, 0.1)';
  ctx.fillRect(0, 0, 960, 540);
  
  ctx.fillStyle = '#ffffff';
  ctx.font = 'bold 56px Arial';
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  
  if (slideNumber) {
    ctx.fillText('Slide ' + slideNumber, 480, 220);
  } else {
    ctx.fillText('Current Slide', 480, 220);
  }
  
  ctx.font = '28px Arial';
  ctx.fillText('Draw your answer below', 480, 300);
  
  ctx.strokeStyle = 'rgba(255, 255, 255, 0.3)';
  ctx.lineWidth = 4;
  ctx.strokeRect(30, 30, 900, 480);
  
  return canvas.toDataURL('image/png');
}

function endActivity() {
  if (currentSessionId) {
    socket.emit('end-session', currentSessionId);
  }
  
  // Clear timer
  if (timerInterval) {
    clearInterval(timerInterval);
    timerInterval = null;
  }
  
  currentSessionId = null;
  document.getElementById('submissionsList').innerHTML = '';
  showSection('startSection');
  showMessage('Activity ended', 'success');
}

function copyUrl() {
  const urlInput = document.getElementById('studentUrl');
  urlInput.select();
  urlInput.setSelectionRange(0, 99999);
  
  try {
    document.execCommand('copy');
    const btn = document.getElementById('copyUrlBtn');
    const originalText = btn.textContent;
    btn.textContent = 'Copied!';
    setTimeout(() => {
      btn.textContent = originalText;
    }, 2000);
  } catch (err) {
    showMessage('Failed to copy URL', 'error');
  }
}

function refreshSubmissions() {
  if (currentSessionId) {
    socket.emit('get-submissions', currentSessionId);
  }
}

function displayAllSubmissions(submissions) {
  const list = document.getElementById('submissionsList');
  list.innerHTML = '';
  
  submissions.forEach(submission => {
    addSubmissionToList(submission);
  });
  
  updateSubmissionCount();
}

function addSubmissionToList(submission) {
  const list = document.getElementById('submissionsList');
  
  const existingItem = list.querySelector(`[data-submission-id="${submission.id}"]`);
  if (existingItem) {
    return;
  }
  
  const item = document.createElement('div');
  item.className = 'submission-item';
  item.setAttribute('data-submission-id', submission.id);
  item.style.cursor = 'pointer';
  
  // Apply hide names if enabled
  var displayName = submission.studentName;
  if (window.activitySettings && window.activitySettings.hideNames) {
    // Count current submissions to get proper number
    var currentCount = document.getElementById('submissionsList').children.length;
    displayName = 'Student ' + (currentCount + 1);
  }
  
  item.innerHTML = `
    <div class="submission-header">
      <strong>${escapeHtml(displayName)}</strong>
      <span class="submission-time">${formatTime(submission.timestamp)}</span>
    </div>
    <img src="${submission.drawingData}" alt="Student drawing" class="submission-preview">
    <button class="btn-small display-btn" data-id="${submission.id}">Display on Slide</button>
  `;
  
  const img = item.querySelector('.submission-preview');
  img.onclick = (e) => {
    e.stopPropagation();
    viewFullSize(submission, displayName);
  };
  
  const displayBtn = item.querySelector('.display-btn');
  displayBtn.onclick = (e) => {
    e.stopPropagation();
    displayOnSlide(submission.id);
  };
  
  list.insertBefore(item, list.firstChild);
}

function updateSubmissionCount() {
  const count = document.getElementById('submissionsList').children.length;
  document.getElementById('submissionCount').textContent = count;
}

function displayOnSlide(submissionId) {
  var submission = null;
  var submissionsList = document.getElementById('submissionsList').children;
  
  for (var i = 0; i < submissionsList.length; i++) {
    if (submissionsList[i].getAttribute('data-submission-id') === submissionId) {
      var img = submissionsList[i].querySelector('.submission-preview');
      if (img) {
        submission = {
          id: submissionId,
          drawingData: img.src
        };
      }
      break;
    }
  }
  
  if (!submission) {
    showMessage('Could not find submission', 'error');
    return;
  }
  
  try {
    var base64Data = submission.drawingData.split(',')[1];
    
    if (!base64Data) {
      showMessage('Invalid image data', 'error');
      return;
    }
    
    Office.context.document.setSelectedDataAsync(
      base64Data,
      {
        coercionType: Office.CoercionType.Image,
        imageLeft: 0,
        imageTop: 0,
        imageWidth: 960,
        imageHeight: 540
      },
      function(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          showMessage('Drawing added to slide successfully!', 'success');
        } else {
          console.error('Error details:', asyncResult.error);
          showMessage('Failed to add drawing: ' + asyncResult.error.message, 'error');
        }
      }
    );
  } catch (error) {
    console.error('Error in displayOnSlide:', error);
    showMessage('Error processing image: ' + error.message, 'error');
  }
}

function viewFullSize(submission, displayName) {
  const modal = document.createElement('div');
  modal.id = 'fullSizeModal';
  modal.style.cssText = `
    position: fixed;
    top: 0;
    left: 0;
    right: 0;
    bottom: 0;
    background: rgba(0,0,0,0.9);
    z-index: 10000;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    padding: 20px;
  `;
  
  const content = document.createElement('div');
  content.style.cssText = `
    background: white;
    border-radius: 12px;
    padding: 20px;
    max-width: 90%;
    max-height: 90%;
    overflow: auto;
    text-align: center;
  `;
  
  content.innerHTML = `
    <h2 style="margin-top: 0;">${escapeHtml(displayName || submission.studentName)}</h2>
    <p style="color: #666;">${formatTime(submission.timestamp)}</p>
    <img src="${submission.drawingData}" style="max-width: 100%; border-radius: 8px; margin: 10px 0;">
    <br>
    <button id="closeModalBtn" style="padding: 10px 30px; background: #dc3545; color: white; border: none; border-radius: 6px; cursor: pointer; margin: 10px;">Close</button>
  `;
  
  modal.appendChild(content);
  document.body.appendChild(modal);
  
  const closeModalFunc = function() {
    if (modal.parentNode) {
      modal.parentNode.removeChild(modal);
    }
  };
  
  document.getElementById('closeModalBtn').onclick = closeModalFunc;
  modal.onclick = function(e) {
    if (e.target === modal) {
      closeModalFunc();
    }
  };
}

function openSubmissionsWindow() {
  if (!currentSessionId) {
    showMessage('No active session', 'error');
    return;
  }
  
  try {
    const viewerUrl = `${SERVER_URL}/viewer?session=${currentSessionId}`;
    
    Office.context.ui.displayDialogAsync(
      viewerUrl,
      { height: 80, width: 70, displayInIframe: false },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const dialog = result.value;
          
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
            console.log('Message from dialog:', arg);
          });
          
          dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
            if (arg.error === 12006) {
              console.log('Dialog closed by user');
            }
          });
        } else {
          console.log('Dialog API failed, using window.open');
          openRegularWindow(viewerUrl);
        }
      }
    );
  } catch (error) {
    console.log('Dialog API not available, using window.open');
    const viewerUrl = `${SERVER_URL}/viewer?session=${currentSessionId}`;
    openRegularWindow(viewerUrl);
  }
}

function openRegularWindow(url) {
  window.open(url, 'SubmissionsViewer', 'width=650,height=400,scrollbars=yes,resizable=yes,toolbar=no,menubar=no,location=no');
  showMessage('Opening submissions viewer...', 'success');
}

function showSection(sectionId) {
  const sections = ['authSection', 'startSection', 'activeSection', 'loadingSection', 'historySection', 'sessionDetailSection'];
  sections.forEach(id => {
    const el = document.getElementById(id);
    if (el) el.style.display = 'none';
  });
  document.getElementById(sectionId).style.display = 'block';
}

function showMessage(message, type) {
  const msgDiv = document.createElement('div');
  msgDiv.className = `message ${type}`;
  msgDiv.textContent = message;
  msgDiv.style.cssText = `
    position: fixed;
    top: 70px;
    left: 50%;
    transform: translateX(-50%);
    padding: 12px 20px;
    background: ${type === 'error' ? '#dc3545' : '#28a745'};
    color: white;
    border-radius: 6px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.2);
    z-index: 1000;
    animation: slideDown 0.3s;
    max-width: 80%;
    text-align: center;
  `;
  
  document.body.appendChild(msgDiv);
  
  setTimeout(() => {
    if (msgDiv.parentNode) {
      msgDiv.parentNode.removeChild(msgDiv);
    }
  }, 3000);
}

function escapeHtml(text) {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

function formatTime(timestamp) {
  const date = new Date(timestamp);
  return date.toLocaleTimeString();
}

window.displayOnSlide = displayOnSlide;