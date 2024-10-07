import { convertDocxToMarkdown } from './docxConverter';

document.getElementById('upload').addEventListener('change', handleFileUpload);

async function handleFileUpload(event) {
    const file = event.target.files[0];
    if (file && file.name.endsWith('.docx')) {
        // Clear existing download links before processing a new file
        clearPreviousDownloadLink();

        // Convert the DOCX to Markdown
        const markdown = await convertDocxToMarkdown(file);

        if (markdown) {
            // Create a new Blob for the Markdown file
            const blob = new Blob([markdown], { type: 'text/markdown' });
            const link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            link.download = 'converted.md';
            link.textContent = 'Download Markdown File';

            // Append the link to the page
            document.getElementById('downloadLink').appendChild(link);
        }
    } else {
        alert('Please upload a valid DOCX file.');
    }
}

// Helper function to clear any previously created download link
function clearPreviousDownloadLink() {
    const downloadLinkContainer = document.getElementById('downloadLink');
    downloadLinkContainer.innerHTML = '';  // Clears any existing content
}
