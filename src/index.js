import { convertDocxToZip } from './docxConverter';

document.getElementById('upload').addEventListener('change', handleFileUpload);

async function handleFileUpload(event) {
    const file = event.target.files[0];

    // Check if the uploaded file is a DOCX file
    if (file && file.name.endsWith('.docx')) {
        // Clear existing download links before processing a new file
        clearPreviousDownloadLink();

        // Show a loading indicator while processing
        const loadingIndicator = document.createElement('p');
        loadingIndicator.textContent = 'Processing... Please wait.';
        document.getElementById('downloadLink').appendChild(loadingIndicator);

        // Convert the DOCX to a ZIP file (Markdown + images)
        try {
            const zipBlob = await convertDocxToZip(file);

            if (zipBlob) {
                // Remove the loading indicator
                // document.getElementById('downloadLink').removeChild(loadingIndicator);

                // Create a new Blob for the ZIP file
                const link = document.createElement('a');
                link.href = URL.createObjectURL(zipBlob);
                link.download = 'converted.zip';
                link.textContent = 'Download ZIP (Markdown + Images)';

                // Append the link to the page
                document.getElementById('downloadLink').replaceChildren(link);
            }
        } catch (error) {
            console.error('Error during file conversion:', error);
            alert('An error occurred while processing the file.');
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
