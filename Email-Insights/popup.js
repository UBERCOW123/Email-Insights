document.addEventListener('DOMContentLoaded', () => {
    // Event Listeners
    document.getElementById('refreshButton').addEventListener('click', refreshData);
    document.getElementById('refreshFromEmpty').addEventListener('click', refreshData);
    document.getElementById('optionsButton').addEventListener('click', () => {
        chrome.runtime.openOptionsPage();
    });
    document.getElementById('generateReport').addEventListener('click', generateReport);
    document.getElementById('sortBy').addEventListener('change', loadStatistics);
    document.getElementById('requestAdminApproval').addEventListener('click', generateAdminRequestEmail);
    document.getElementById('buyMeACoffee').addEventListener('click', () => {
        window.open('https://www.buymeacoffee.com/example', '_blank'); // Replace with your actual link
    });

    // Initial load
    loadStatistics();
});

function refreshData() {
    showLoading(true);
    document.getElementById('adminConsent').classList.add('hidden'); // Hide error on refresh
    chrome.runtime.sendMessage({ action: 'analyzeEmails' }, () => {
        // Add a slight delay to allow storage to update before reloading stats
        setTimeout(() => {
            loadStatistics();
        }, 2000); 
    });
}

function showLoading(show) {
    const loading = document.getElementById('loading');
    const statsContainer = document.getElementById('statsContainer');
    const summaryStats = document.getElementById('summaryStats');
    const emptyState = document.getElementById('emptyState');
    const adminConsent = document.getElementById('adminConsent');

    if (show) {
        loading.classList.remove('hidden');
        statsContainer.classList.add('hidden');
        summaryStats.classList.add('hidden');
        emptyState.classList.add('hidden');
        adminConsent.classList.add('hidden');
    } else {
        loading.classList.add('hidden');
    }
}

function loadStatistics() {
    const container = document.getElementById('statsContainer');
    const summaryStats = document.getElementById('summaryStats');
    const emptyState = document.getElementById('emptyState');
    const adminConsent = document.getElementById('adminConsent');

    showLoading(true);

    chrome.storage.local.get(['emailStats', 'settings', 'adminConsentRequired'], (data) => {
        showLoading(false);
        
        // Handle Admin Consent Error first
        if (data.adminConsentRequired) {
            container.classList.add('hidden');
            summaryStats.classList.add('hidden');
            emptyState.classList.add('hidden');
            adminConsent.classList.remove('hidden');
            return;
        }

        if (!data.emailStats || Object.keys(data.emailStats).length === 0) {
            container.classList.add('hidden');
            summaryStats.classList.add('hidden');
            emptyState.classList.remove('hidden');
            return;
        }

        const defaultSettings = {
            shameThreshold: 30,
            ignoreThreshold: 10
        };
        const settings = data.settings || defaultSettings;
        const stats = data.emailStats;
        const sortBy = document.getElementById('sortBy').value;
        
        emptyState.classList.add('hidden');
        container.classList.remove('hidden');
        summaryStats.classList.remove('hidden');

        // Calculate summary statistics
        displaySummaryStats(stats);

        // Sort and display contact cards
        const sortedContacts = sortContacts(stats, sortBy);
        displayStats(sortedContacts, settings);
    });
}

function generateAdminRequestEmail() {
    const subject = "Request for App Approval: Outlook Insights Dashboard";
    const body = `
Hi IT Administrator,

I would like to request approval for a Chrome Extension called "Outlook Insights Dashboard".

This extension helps me analyze my email communication patterns locally in my browser to improve my productivity.

Key points for your review:
- It is privacy-focused and uses the most limited Microsoft Graph permission possible ("Mail.ReadBasic").
- This permission allows it to read only basic email metadata (like sender, subject, and date), but NOT the email body content.
- All data processing happens locally on my computer. No data is sent to any external servers.
- The application's Client ID is: 6dd93ff2-c48b-4c69-8555-605b5ab500ee

To approve this application, you can grant tenant-wide admin consent for the Client ID mentioned above.

Thank you for your consideration.
    `;
    
    const mailtoLink = `mailto:?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body.trim())}`;
    window.open(mailtoLink, '_blank');
}


function displaySummaryStats(stats) {
    let totalContacts = 0;
    let totalSent = 0;
    let totalReceived = 0;
    let totalIgnored = 0;
    let totalResponseTimes = [];

    for (const contact in stats) {
        totalContacts++;
        totalSent += stats[contact].sent || 0;
        totalReceived += stats[contact].received || 0;
        totalIgnored += stats[contact].ignored || 0;

        const responseTime = parseResponseTime(stats[contact].avgResponseTime);
        if (responseTime !== null) {
            totalResponseTimes.push(responseTime);
        }
    }

    const totalEmails = totalSent + totalReceived;
    const avgResponseTime = totalResponseTimes.length > 0
        ? formatResponseTime(totalResponseTimes.reduce((a, b) => a + b, 0) / totalResponseTimes.length)
        : '-';

    document.getElementById('totalContacts').textContent = totalContacts;
    document.getElementById('totalEmails').textContent = totalEmails;
    document.getElementById('totalIgnored').textContent = totalIgnored;
    document.getElementById('avgResponseTime').textContent = avgResponseTime;
}

function sortContacts(stats, sortBy) {
    const contacts = Object.entries(stats).map(([email, data]) => ({
        email,
        ...data
    }));

    switch (sortBy) {
        case 'ignored':
            return contacts.sort((a, b) => (b.ignored || 0) - (a.ignored || 0));
        case 'responseTime':
            return contacts.sort((a, b) => {
                const timeA = parseResponseTime(a.avgResponseTime);
                const timeB = parseResponseTime(b.avgResponseTime);
                if (timeA === null) return 1;
                if (timeB === null) return -1;
                return timeB - timeA;
            });
        case 'volume':
            return contacts.sort((a, b) => {
                const volumeA = (a.sent || 0) + (a.received || 0);
                const volumeB = (b.sent || 0) + (b.received || 0);
                return volumeB - volumeA;
            });
        case 'name':
        default:
            return contacts.sort((a, b) => a.email.localeCompare(b.email));
    }
}

function displayStats(contacts, settings) {
    const container = document.getElementById('statsContainer');
    container.innerHTML = '';

    contacts.forEach(contact => {
        const card = createContactCard(contact, settings);
        container.appendChild(card);
    });
}

function createContactCard(contact, settings) {
    const card = document.createElement('div');
    card.className = 'contact-card';

    const totalSent = contact.sent || 0;
    const totalIgnored = contact.ignored || 0;
    const repliedTo = contact.repliedTo || 0;
    
    const replyRate = totalSent > 0 ? (repliedTo / totalSent * 100) : 100;

    const isShame = (totalSent > 5 && replyRate < settings.shameThreshold) || totalIgnored >= settings.ignoreThreshold;
    if (isShame) {
        card.classList.add('shame');
    }

    card.innerHTML = `
        <div class="contact-header">
            <h3>${formatEmail(contact.email)}</h3>
            ${isShame ? '<span class="shame-badge">Low Reply Rate</span>' : ''}
        </div>
        <div class="stats-grid">
            <div class="stat-item">
                <div class="stat-item-label">Sent To</div>
                <div class="stat-item-value">${totalSent}</div>
            </div>
            <div class="stat-item">
                <div class="stat-item-label">Received From</div>
                <div class="stat-item-value">${contact.received || 0}</div>
            </div>
            <div class="stat-item">
                <div class="stat-item-label">Replied To You</div>
                <div class="stat-item-value">${repliedTo}</div>
            </div>
            <div class="stat-item">
                <div class="stat-item-label">Ignored You</div>
                <div class="stat-item-value ${totalIgnored >= settings.ignoreThreshold ? 'bad' : ''}">${totalIgnored}</div>
            </div>
        </div>
        <div class="progress-bar">
            <div class="progress-label">
                <span>Reply Rate</span>
                <span>${Math.round(replyRate)}%</span>
            </div>
            <div class="progress-track">
                <div class="progress-fill ${replyRate < settings.shameThreshold ? 'low' : ''}" style="width: ${replyRate}%"></div>
            </div>
        </div>
        <div class="stat-item" style="margin-top: 10px;">
            <div class="stat-item-label">Avg Response Time</div>
            <div class="stat-item-value">${contact.avgResponseTime || '-'}</div>
        </div>
    `;

    return card;
}

function formatEmail(email) {
    // Truncate long emails
    if (email.length > 30) {
        return email.substring(0, 27) + '...';
    }
    return email;
}

function parseResponseTime(timeStr) {
    if (!timeStr || timeStr === '-') return null;

    const hourMatch = timeStr.match(/([\d.]+)\s*h/);
    const dayMatch = timeStr.match(/([\d.]+)\s*d/);
    const minMatch = timeStr.match(/([\d.]+)\s*m/);

    let minutes = 0;
    if (dayMatch) minutes += parseFloat(dayMatch[1]) * 24 * 60;
    if (hourMatch) minutes += parseFloat(hourMatch[1]) * 60;
    if (minMatch) minutes += parseFloat(minMatch[1]);

    return minutes || null;
}

function formatResponseTime(minutes) {
    if (minutes < 60) {
        return `${Math.round(minutes)}m`;
    } else if (minutes < 1440) {
        return `${(minutes / 60).toFixed(1)}h`;
    } else {
        return `${(minutes / 1440).toFixed(1)}d`;
    }
}

function generateReport() {
    chrome.storage.local.get('emailStats', (data) => {
        if (!data.emailStats || Object.keys(data.emailStats).length === 0) {
            alert('No data available to generate report. Please refresh data first.');
            return;
        }

        const report = createHTMLReport(data.emailStats);
        const blob = new Blob([report], { type: 'text/html' });
        const url = URL.createObjectURL(blob);

        const a = document.createElement('a');
        a.href = url;
        a.download = `outlook-insights-report-${new Date().toISOString().split('T')[0]}.html`;
        a.click();

        URL.revokeObjectURL(url);
    });
}

function createHTMLReport(stats) {
    const date = new Date().toLocaleDateString();
    const contacts = Object.entries(stats).map(([email, data]) => ({
        email,
        ...data
    })).sort((a, b) => (b.ignored || 0) - (a.ignored || 0));

    let contactRows = '';
    contacts.forEach(contact => {
        const replyRate = (contact.sent || 0) > 0
            ? Math.round(((contact.repliedTo || 0) / contact.sent) * 100)
            : 0;

        contactRows += `
            <tr>
                <td>${contact.email}</td>
                <td>${contact.sent || 0}</td>
                <td>${contact.received || 0}</td>
                <td>${contact.repliedTo || 0}</td>
                <td style="color: ${contact.ignored > 5 ? '#ff6b6b' : '#333'}; font-weight: bold;">${contact.ignored || 0}</td>
                <td>${replyRate}%</td>
                <td>${contact.avgResponseTime || '-'}</td>
            </tr>
        `;
    });

    return `
<!DOCTYPE html>
<html>
<head>
    <title>Outlook Insights Report - ${date}</title>
    <style>
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            max-width: 1200px;
            margin: 40px auto;
            padding: 40px;
            background: #f5f5f5;
        }
        .report {
            background: white;
            border-radius: 12px;
            padding: 40px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.1);
        }
        h1 {
            color: #667eea;
            border-bottom: 3px solid #667eea;
            padding-bottom: 20px;
            margin-bottom: 30px;
        }
        .meta {
            color: #666;
            margin-bottom: 30px;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
        }
        th, td {
            padding: 12px;
            text-align: left;
            border-bottom: 1px solid #e0e0e0;
        }
        th {
            background: #667eea;
            color: white;
            font-weight: 600;
        }
        tr:hover {
            background: #f8f9fa;
        }
        .footer {
            margin-top: 40px;
            padding-top: 20px;
            border-top: 1px solid #e0e0e0;
            color: #666;
            font-size: 14px;
        }
    </style>
</head>
<body>
    <div class="report">
        <h1>Outlook Insights Report</h1>
        <div class="meta">
            <strong>Generated:</strong> ${date}<br>
            <strong>Total Contacts:</strong> ${contacts.length}
        </div>
        <table>
            <thead>
                <tr>
                    <th>Contact</th>
                    <th>Sent To</th>
                    <th>Received From</th>
                    <th>Replied To You</th>
                    <th>Ignored You</th>
                    <th>Reply Rate</th>
                    <th>Avg Response Time</th>
                </tr>
            </thead>
            <tbody>
                ${contactRows}
            </tbody>
        </table>
        <div class="footer">
            <p>Generated by Outlook Insights Dashboard</p>
        </div>
    </div>
</body>
</html>
    `;
}