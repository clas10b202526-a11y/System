// CONFIGURATION
// -------------------------------------------------------
// Replace this string with the exact name of your file inside the folder
const presentationFile = "presentation.pptx"; 
const presentationTitle = "Q4 Strategy Report"; // Shown in sidebar
// -------------------------------------------------------

document.addEventListener('DOMContentLoaded', () => {
    
    // 1. Update text info
    document.getElementById('doc-title').textContent = presentationTitle;
    document.getElementById('download-btn').href = presentationFile;

    // 2. check if we are offline (local file system)
    // The Microsoft Viewer requires a live HTTP/HTTPS URL.
    const isLocal = window.location.protocol === 'file:';
    
    if (isLocal) {
        // Show error message if user tries to open directly from folder
        document.getElementById('loader').classList.add('hidden');
        document.getElementById('error-msg').classList.remove('hidden');
        return;
    }

    // 3. Construct the Full URL
    // This grabs the current website address (e.g., yourdomain.com/folder/)
    // and appends the filename.
    const currentPath = window.location.href.substring(0, window.location.href.lastIndexOf('/'));
    const fullFileUrl = `${currentPath}/${presentationFile}`;

    // 4. Load into Viewer
    const viewerBase = "https://view.officeapps.live.com/op/embed.aspx?src=";
    const finalUrl = viewerBase + encodeURIComponent(fullFileUrl);
    
    const iframe = document.getElementById('ppt-frame');
    const loader = document.getElementById('loader');

    iframe.src = finalUrl;

    iframe.onload = () => {
        loader.classList.add('hidden');
    };
    
    // Fallback timeout in case the 'load' event doesn't fire perfectly
    setTimeout(() => {
        loader.classList.add('hidden');
    }, 4000);
});
