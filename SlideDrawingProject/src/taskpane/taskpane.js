/* global Office, io */

let socket;
let currentSessionId = null;
let uploadedSlideImage = null;
let timerInterval = null;
const SERVER_URL = 'http://localhost:3001';

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    setupEventListeners();
    connectToServer();
  }
});

function setupEventListeners() {
  document.getElementById('startButton').onclick = startActivity;
  document.getElementById('endButton').onclick = endActivity;
  document.getElementById('copyUrlBtn').onclick = copyUrl;
  document.getElementById('refreshBtn').onclick = refreshSubmissions;
  document.getElementById('viewAllBtn').onclick = openSubmissionsWindow;
  document.getElementById('slideImageUpload').onchange = handleImageUpload;
  document.getElementById('removeImageBtn').onclick = removeImage;
  
  // Timer option toggle
  var timerRadios = document.getElementsByName('timerOption');
  for (var i = 0; i < timerRadios.length; i++) {
    timerRadios[i].onchange = function() {
      var timerInput = document.getElementById('timerMinutes');
      timerInput.disabled = this.value === 'unlimited';
    };
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
    
    // Create session with settings - INCLUDING TIMER!
    console.log('Creating session via Socket.io...');
    socket.emit('create-session', { 
      slideImage: slideImage,
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
      Office.context.document.getSelectedDataAsync(
        Office.CoercionType.SlideRange,
        async (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve(createStyledPlaceholder());
          } else {
            resolve(createStyledPlaceholder());
          }
        }
      );
    } catch (error) {
      console.error('Error:', error);
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
  document.getElementById('startSection').style.display = 'none';
  document.getElementById('activeSection').style.display = 'none';
  document.getElementById('loadingSection').style.display = 'none';
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