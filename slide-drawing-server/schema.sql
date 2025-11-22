-- Create database
CREATE DATABASE IF NOT EXISTS slide_drawing_db;
USE slide_drawing_db;

-- Sessions table
CREATE TABLE IF NOT EXISTS sessions (
  id VARCHAR(20) PRIMARY KEY,
  slide_image LONGTEXT NOT NULL,
  teacher_socket_id VARCHAR(100),
  timer_duration INT DEFAULT NULL,
  hide_names BOOLEAN DEFAULT FALSE,
  allow_resubmit BOOLEAN DEFAULT TRUE,
  start_time BIGINT NOT NULL,
  timer_end BIGINT DEFAULT NULL,
  closed BOOLEAN DEFAULT FALSE,
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  INDEX idx_closed (closed),
  INDEX idx_created_at (created_at)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- Drawings/submissions table
CREATE TABLE IF NOT EXISTS drawings (
  id VARCHAR(20) PRIMARY KEY,
  session_id VARCHAR(20) NOT NULL,
  student_name VARCHAR(255) NOT NULL,
  drawing_data LONGTEXT NOT NULL,
  timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  FOREIGN KEY (session_id) REFERENCES sessions(id) ON DELETE CASCADE,
  INDEX idx_session_id (session_id),
  INDEX idx_student_name (student_name),
  UNIQUE KEY unique_student_per_session (session_id, student_name)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
