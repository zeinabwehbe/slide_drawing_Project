// Select the canvas element and get its 2D context
const canvas = document.getElementById('drawingBoard');
const ctx = canvas.getContext('2d');

// Define the size of the canvas
// ClassPoint slides are typically 16:9 aspect ratio
canvas.width = 800;
canvas.height = 450; 

let isDrawing = false;
let currentTool = 'pen';
let currentColor = '#000000';
const penWidth = 3;
const eraserWidth = 10;

// Event listeners for mouse actions
canvas.addEventListener('mousedown', startDrawing);
canvas.addEventListener('mousemove', draw);
canvas.addEventListener('mouseup', stopDrawing);
canvas.addEventListener('mouseout', stopDrawing); // Stop drawing if cursor leaves canvas

// Tool and Color Selector listeners
document.getElementById('penTool').addEventListener('click', () => { currentTool = 'pen'; });
document.getElementById('eraserTool').addEventListener('click', () => { currentTool = 'eraser'; });
document.getElementById('colorPicker').addEventListener('change', (e) => { currentColor = e.target.value; });

function startDrawing(e) {
    isDrawing = true;
    // Start drawing a single dot immediately upon click
    draw(e); 
}

function stopDrawing() {
    isDrawing = false;
    ctx.beginPath(); // Reset the current path when the mouse is released
}

function draw(e) {
    if (!isDrawing) return;

    // Calculate mouse position relative to the canvas
    const rect = canvas.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;

    ctx.lineWidth = (currentTool === 'pen') ? penWidth : eraserWidth;
    ctx.lineCap = 'round';
    ctx.strokeStyle = (currentTool === 'pen') ? currentColor : '#FFFFFF'; // Eraser draws white

    ctx.lineTo(x, y);
    ctx.stroke();
    ctx.beginPath(); // Start a new path from the current point for smooth drawing
    ctx.moveTo(x, y);
}
