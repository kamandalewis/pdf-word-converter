// Progress bar functionality
function updateProgress(percent, message) {
    const progressFill = document.getElementById('progressFill');
    const progressText = document.getElementById('progressText');
    
    progressFill.style.width = percent + '%';
    progressText.textContent = message;
}

// Handle file upload
document.getElementById('pdfInput').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (file) {
        convertPDF(file);
    }
});

// Drag and drop functionality
const uploadArea = document.getElementById('uploadArea');
uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadArea.style.borderColor = 'rgba(255, 255, 255, 0.6)';
    uploadArea.style.transform = 'scale(1.02)';
});

uploadArea.addEventListener('dragleave', () => {
    uploadArea.style.borderColor = 'rgba(255, 255, 255, 0.3)';
    uploadArea.style.transform = 'scale(1)';
});

uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.style.borderColor = 'rgba(255, 255, 255, 0.3)';
    uploadArea.style.transform = 'scale(1)';
    
    const file = e.dataTransfer.files[0];
    if (file && file.type === 'application/pdf') {
        document.getElementById('pdfInput').files = e.dataTransfer.files;
        convertPDF(file);
    }
});

// Main conversion function using PDF.js
async function convertPDF(file) {
    document.getElementById('uploadArea').style.display = 'none';
    document.getElementById('progressContainer').style.display = 'block';
    updateProgress(10, 'Loading PDF...');
    
    try {
        const arrayBuffer = await file.arrayBuffer();
        updateProgress(30, 'Reading document...');
        
        // Load PDF using PDF.js
        const loadingTask = pdfjsLib.getDocument({data: arrayBuffer});
        const pdf = await loadingTask.promise;
        
        updateProgress(50, 'Extracting text...');
        
        let fullText = '';
        const numPages = pdf.numPages;
        
        // Extract text from each page
        for (let i = 1; i <= numPages; i++) {
            const page = await pdf.getPage(i);
            const textContent = await page.getTextContent();
            const pageText = textContent.items.map(item => item.str).join(' ');
            fullText += `${pageText}\n\n`;
            
            updateProgress(50 + (40 * i / numPages), `Processing page ${i}/${numPages}...`);
        }
        
        updateProgress(95, 'Creating Word document...');
        await createWordDocument(file, fullText);
        
        updateProgress(100, 'Complete!');
        await new Promise(resolve => setTimeout(resolve, 300));
        
    } catch (error) {
        console.error('Conversion error:', error);
        alert('Error reading PDF. Please ensure the file is not corrupted or password-protected.');
        document.getElementById('progressContainer').style.display = 'none';
        document.getElementById('uploadArea').style.display = 'block';
    }
}

// Create actual .docx file
async function createWordDocument(file, extractedText) {
    try {
        // Create a new document using docx library
        const doc = new docx.Document({
            sections: [{
                properties: {},
                children: extractedText.split('\n').map(line => 
                    new docx.Paragraph({
                        text: line,
                        spacing: {
                            after: 100,
                        },
                    })
                ),
            }],
        });
        
        // Generate the document as a blob
        const blob = await docx.Packer.toBlob(doc);
        const url = URL.createObjectURL(blob);
        
        const downloadLink = document.getElementById('downloadLink');
        downloadLink.href = url;
        downloadLink.download = file.name.replace('.pdf', '.docx');
        downloadLink.textContent = '📥 Download Word Document (.docx)';
        
        // Also provide copy-to-clipboard option
        document.getElementById('textPreview').textContent = extractedText;
        document.getElementById('copyBtn').onclick = () => {
            navigator.clipboard.writeText(extractedText);
            document.getElementById('copyBtn').textContent = '✓ Copied!';
            setTimeout(() => {
                document.getElementById('copyBtn').textContent = '📋 Copy Text';
            }, 2000);
        };
        
        // Show results
        document.getElementById('progressContainer').style.display = 'none';
        document.getElementById('resultArea').style.display = 'block';
        
        setTimeout(() => {
            document.getElementById('ratingSection').style.opacity = '1';
        }, 1000);
        
    } catch (error) {
        console.error('Word creation error:', error);
        alert('Error creating Word document. Please try again.');
        document.getElementById('progressContainer').style.display = 'none';
        document.getElementById('uploadArea').style.display = 'block';
    }
}

// Rating functionality
let rating = 0;
function initStars() {
    const starsContainer = document.getElementById('stars');
    starsContainer.innerHTML = '';
    for (let i = 1; i <= 5; i++) {
        const span = document.createElement('span');
        span.textContent = '☆';
        span.style.cursor = 'pointer';
        span.style.fontSize = '24px';
        span.dataset.rating = i;
        span.addEventListener('click', function() {
            rating = i;
            updateStars(i);
        });
        span.addEventListener('mouseenter', function() {
            updateStars(i, true);
        });
        starsContainer.addEventListener('mouseleave', function() {
            updateStars(rating);
        });
        starsContainer.appendChild(span);
    }
}

function updateStars(count, isHover = false) {
    const starElements = document.querySelectorAll('#stars span');
    starElements.forEach((star, index) => {
        star.textContent = index < count ? '⭐' : '☆';
    });
}

// Share tool function
function shareTool() {
    if (navigator.share) {
        navigator.share({
            title: 'Free PDF to Word Converter',
            text: 'Convert PDF to Word instantly - no signup required!',
            url: window.location.href
        });
    } else {
        navigator.clipboard.writeText(window.location.href);
        alert('Link copied to clipboard!');
    }
}

// Convert another file
function convertAnother() {
    document.getElementById('resultArea').style.display = 'none';
    document.getElementById('uploadArea').style.display = 'block';
    document.getElementById('pdfInput').value = '';
    document.getElementById('textPreview').value = '';
    rating = 0;
}

// Initialize when page loads
document.addEventListener('DOMContentLoaded', initStars);