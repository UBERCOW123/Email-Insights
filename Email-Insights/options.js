// Enhanced options with additional customization

document.addEventListener('DOMContentLoaded', () => {
    restoreOptions();
    setupEventListeners();
});

function setupEventListeners() {
    document.getElementById('save').addEventListener('click', saveOptions);
    document.getElementById('cancel').addEventListener('click', () => window.close());
    document.getElementById('clearData').addEventListener('click', clearAllData);

    // Update slider value displays
    const shameThreshold = document.getElementById('shameThreshold');
    const ignoreThreshold = document.getElementById('ignoreThreshold');

    shameThreshold.addEventListener('input', (e) => {
        document.getElementById('shameThresholdValue').textContent = e.target.value + '%';
    });

    ignoreThreshold.addEventListener('input', (e) => {
        document.getElementById('ignoreThresholdValue').textContent = e.target.value;
    });
}

function saveOptions() {
    const settings = {
        analysisPeriod: parseInt(document.getElementById('analysisPeriod').value),
        filterCalendarInvites: document.getElementById('filterCalendarInvites').checked,
        filterOutOfOffice: document.getElementById('filterOutOfOffice').checked,
        filterGroupEmails: document.getElementById('filterGroupEmails').checked,
        useBusinessHours: document.getElementById('useBusinessHours').checked,
        shameThreshold: parseInt(document.getElementById('shameThreshold').value),
        ignoreThreshold: parseInt(document.getElementById('ignoreThreshold').value)
    };

    chrome.storage.local.set({ settings: settings }, () => {
        showStatus('Settings saved successfully!', 'success');
    });
}

function restoreOptions() {
    const defaultSettings = {
        analysisPeriod: 90,
        filterCalendarInvites: true,
        filterOutOfOffice: true,
        filterGroupEmails: true,
        useBusinessHours: true,
        shameThreshold: 30,
        ignoreThreshold: 10
    };

    chrome.storage.local.get({ settings: defaultSettings }, (data) => {
        const settings = data.settings;

        document.getElementById('analysisPeriod').value = settings.analysisPeriod;
        document.getElementById('filterCalendarInvites').checked = settings.filterCalendarInvites;
        document.getElementById('filterOutOfOffice').checked = settings.filterOutOfOffice;
        document.getElementById('filterGroupEmails').checked = settings.filterGroupEmails;
        document.getElementById('useBusinessHours').checked = settings.useBusinessHours;
        
        document.getElementById('shameThreshold').value = settings.shameThreshold;
        document.getElementById('shameThresholdValue').textContent = settings.shameThreshold + '%';
        
        document.getElementById('ignoreThreshold').value = settings.ignoreThreshold;
        document.getElementById('ignoreThresholdValue').textContent = settings.ignoreThreshold;
    });
}

function clearAllData() {
    if (confirm('Are you sure you want to clear all stored email statistics? This action cannot be undone.')) {
        chrome.storage.local.remove('emailStats', () => {
            showStatus('All stored data has been cleared.', 'success');
            
            // Optionally trigger a re-analysis
            setTimeout(() => {
                if (confirm('Would you like to re-analyze your emails now?')) {
                    chrome.runtime.sendMessage({ action: 'analyzeEmails' });
                }
            }, 1000);
        });
    }
}

function showStatus(message, type) {
    const status = document.getElementById('status');
    status.textContent = message;
    status.className = type;
    status.classList.remove('hidden');

    setTimeout(() => {
        status.classList.add('hidden');
    }, 3000);
}