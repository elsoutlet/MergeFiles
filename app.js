// Core app logic for Excel File Merger Tool

let selectedFiles = [];
let buttonDisabled = true;

const filesInput = document.getElementById('files');
const masterFileInput = document.getElementById('masterFile');
const mergeBtn = document.getElementById('mergeBtn');
const masterBtn = document.getElementById('masterBtn');
const mergeSpinner = document.getElementById('mergeSpinner');
const masterSpinner = document.getElementById('masterSpinner');
const statusDiv = document.getElementById('status');

function setStatus(message, type = 'info') {
    statusDiv.textContent = message;
    statusDiv.className = `status status-${type}`;
}

function setLoading(btnId, spinnerId, isLoading) {
    const btn = document.getElementById(btnId);
    const spinner = document.getElementById(spinnerId);
    if (btn && spinner) {
        btn.disabled = isLoading;
        spinner.style.display = isLoading ? 'inline-block' : 'none';
    }
}

function updateButtonState() {
    buttonDisabled = selectedFiles.length === 0;
    mergeBtn.disabled = buttonDisabled;
    masterBtn.disabled = buttonDisabled;
}

function handleFileChange(event) {
    const input = event.target;
    selectedFiles = input.files ? Array.from(input.files) : [];
    updateButtonState();
    // Optionally, update file list UI
    const fileListDiv = document.getElementById('fileList');
    if (fileListDiv) {
        fileListDiv.innerHTML = selectedFiles.map(f => `<div>${f.name}</div>`).join('');
    }
}

async function downloadMergedFiles() {
    if (selectedFiles.length === 0) {
        setStatus('Please select files first', 'error');
        return;
    }
    setLoading('mergeBtn', 'mergeSpinner', true);
    setStatus('Processing files...', 'info');
    try {
        const result = await window.processFiles.merge(selectedFiles);
        if (!result || !result.length) {
            setStatus('No data could be merged from the uploaded files', 'error');
            return;
        }
        let downloadCount = 0;
        result.forEach(([csvContent, fileName], index) => {
            const name = fileName || `merged_file_${index + 1}`;
            const blob = new Blob([csvContent], { type: 'text/csv' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `${name}.csv`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
            downloadCount++;
        });
        setStatus(`Successfully downloaded ${downloadCount} merged files`, 'success');
    } catch (e) {
        console.error(e);
        setStatus('Failed to process files', 'error');
    } finally {
        setLoading('mergeBtn', 'mergeSpinner', false);
    }
}

async function mergeAndDownloadMaster() {
    let masterFile = null;
    if (masterFileInput && masterFileInput.files && masterFileInput.files.length > 0) {
        masterFile = masterFileInput.files[0];
    }
    if (selectedFiles.length === 0) {
        setStatus('Please select files first', 'error');
        return;
    }
    setLoading('masterBtn', 'masterSpinner', true);
    setStatus('Processing files for master merge...', 'info');
    try {
        const wb = await window.processFiles.combineMergedFiles(selectedFiles, masterFile);
        if (!wb) {
            setStatus('No data in combined file', 'error');
            return;
        }
        // Write workbook to blob and trigger download
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'merged_master.xlsx';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        setStatus('Master Excel file downloaded', 'success');
    } catch (e) {
        console.error('Error saving file:', e);
        setStatus('Failed to save the file', 'error');
    } finally {
        setLoading('masterBtn', 'masterSpinner', false);
    }
}

// Attach event listeners
document.addEventListener('DOMContentLoaded', function () {
    if (filesInput) filesInput.addEventListener('change', handleFileChange);
    if (mergeBtn) mergeBtn.addEventListener('click', downloadMergedFiles);
    if (masterBtn) masterBtn.addEventListener('click', mergeAndDownloadMaster);
    updateButtonState();
    setLoading('mergeBtn', 'mergeSpinner', false);
    setLoading('masterBtn', 'masterSpinner', false);
}); 