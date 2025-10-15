// Enhanced email analysis with Microsoft Graph API integration

chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    if (request.action === 'analyzeEmails') {
        analyzeEmailData();
        sendResponse({ success: true });
        return true;
    }
});

async function analyzeEmailData() {
    try {
        // Clear any previous admin consent errors before a new attempt
        await chrome.storage.local.remove('adminConsentRequired');

        // 1. Get user settings
        const { settings } = await chrome.storage.local.get('settings');
        const config = settings || getDefaultSettings();

        // 2. Get OAuth token
        const token = await getAuthToken();

        if (!token) {
            console.error('Failed to get authentication token. This may require admin consent.');
            // The getAuthToken function will set the adminConsentRequired flag if needed
            return;
        }

        // 3. Fetch emails from Microsoft Graph API using the configured date range
        const emails = await fetchEmails(token, config);

        // 4. Process emails according to settings
        const stats = processEmails(emails, config);

        // 5. Save results
        await chrome.storage.local.set({ emailStats: stats });
        console.log('Email analysis complete. Analyzed', Object.keys(stats).length, 'contacts.');

    } catch (error) {
        console.error('Error analyzing emails:', error);
        // Fallback to mock data for development
        useMockData();
    }
}

function getDefaultSettings() {
    return {
        analysisPeriod: 90,
        filterCalendarInvites: true,
        filterOutOfOffice: true,
        filterGroupEmails: true,
        useBusinessHours: true
    };
}

async function getAuthToken() {
    return new Promise((resolve) => {
        const clientId = '6dd93ff2-c48b-4c69-8555-605b5ab500ee'; // From your manifest
        const scopes = ['https://graph.microsoft.com/Mail.ReadBasic', 'https://graph.microsoft.com/User.Read'];
        const redirectUri = chrome.identity.getRedirectURL();

        const authUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?` +
            `client_id=${encodeURIComponent(clientId)}` +
            `&response_type=token` +
            `&redirect_uri=${encodeURIComponent(redirectUri)}` +
            `&scope=${encodeURIComponent(scopes.join(' '))}` +
            `&response_mode=fragment` +
            `&state=1234` +
            `&nonce=678910`;

        chrome.identity.launchWebAuthFlow({
            url: authUrl,
            interactive: true
        }, (redirectUrl) => {
            if (chrome.runtime.lastError || !redirectUrl) {
                console.error('Auth error:', chrome.runtime.lastError?.message);
                resolve(null);
                return;
            }
            
            // Check for admin consent required error
            if (redirectUrl.includes('error=interaction_required') && redirectUrl.includes('AADSTS65001')) {
                console.error('Admin consent is required.');
                chrome.storage.local.set({ adminConsentRequired: true });
                resolve(null);
                return;
            }

            // Parse the access_token from the redirect URL fragment
            const hash = redirectUrl.split('#')[1];
            const params = new URLSearchParams(hash);
            const token = params.get('access_token');

            if (token) {
                resolve(token);
            } else {
                console.error('No token in redirect URL:', redirectUrl);
                resolve(null);
            }
        });
    });
}


async function fetchEmails(token, config) {
    const headers = {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
    };

    // Note: All-time analysis is handled by fetching with no date filter.
    const analysisDays = config.analysisPeriod;
    let dateFilterString = '';
    if (analysisDays !== -1) { // -1 represents "All-Time"
        const dateToFilter = new Date();
        dateToFilter.setDate(dateToFilter.getDate() - analysisDays);
        dateFilterString = `$filter=receivedDateTime ge ${dateToFilter.toISOString()}&`;
    }

    console.log(`Fetching emails with period: ${analysisDays === -1 ? 'All-Time' : `Last ${analysisDays} days`}...`);
    const selectFields = '$select=subject,from,toRecipients,ccRecipients,receivedDateTime,conversationId,isRead,sender';

    // Get sent emails
    const sentResponse = await fetch(
        `https://graph.microsoft.com/v1.0/me/mailFolders/sentitems/messages?${dateFilterString}${selectFields}&$top=1000`,
        { headers }
    );

    // Get inbox emails
    const inboxResponse = await fetch(
        `https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?${dateFilterString}${selectFields}&$top=1000`,
        { headers }
    );

    const sentData = await sentResponse.json();
    const inboxData = await inboxResponse.json();

    return {
        sent: sentData.value || [],
        received: inboxData.value || []
    };
}

function shouldFilterEmail(email, config) {
    const subject = email.subject?.toLowerCase() || '';

    // Filter calendar invites
    if (config.filterCalendarInvites) {
        if (subject.startsWith('accepted:') || subject.startsWith('declined:') || subject.startsWith('tentative:') || 
            subject.includes('invitation:') || subject.includes('meeting request')) {
            return true;
        }
    }

    // Filter out of office / auto-replies based on subject (since headers require Mail.Read)
    if (config.filterOutOfOffice) {
        const autoReplySubjects = [
            'out of office', 'ooo', 'automatic reply', 'autoreply',
            'away from the office', 'out of the office', 'undeliverable:',
            'delivery status notification (failure)', 'message could not be delivered'
        ];
        if (autoReplySubjects.some(s => subject.includes(s))) {
            return true;
        }
    }

    return false;
}

function processEmails(emails, config) {
    const contactStats = {};
    const conversations = {};

    // Process sent emails
    emails.sent.forEach(email => {
        if (shouldFilterEmail(email, config)) return;

        const recipients = extractRecipients(email, config.filterGroupEmails);
        recipients.forEach(recipient => {
            if (!contactStats[recipient]) {
                contactStats[recipient] = initContactStats();
            }
            contactStats[recipient].sent++;

            // Track conversation for response time calculation
            if (email.conversationId) {
                if (!conversations[email.conversationId]) {
                    conversations[email.conversationId] = [];
                }
                conversations[email.conversationId].push({
                    type: 'sent',
                    timestamp: new Date(email.receivedDateTime),
                    recipient: recipient
                });
            }
        });
    });

    // Process received emails
    emails.received.forEach(email => {
        if (shouldFilterEmail(email, config)) return;

        const sender = extractSender(email);
        if (!sender) return;

        if (!contactStats[sender]) {
            contactStats[sender] = initContactStats();
        }
        contactStats[sender].received++;

        // Track conversation
        if (email.conversationId) {
            if (!conversations[email.conversationId]) {
                conversations[email.conversationId] = [];
            }
            conversations[email.conversationId].push({
                type: 'received',
                timestamp: new Date(email.receivedDateTime),
                sender: sender
            });
        }
    });

    // Calculate response metrics
    calculateResponseMetrics(contactStats, conversations, config);

    return contactStats;
}
function extractRecipients(email, separateGroups) {
    const recipients = [];
    const allRecipients = [
        ...(email.toRecipients || []),
        ...(email.ccRecipients || [])
    ];

    if (separateGroups) {
        // Return individual recipients
        allRecipients.forEach(recipient => {
            if (recipient.emailAddress?.address) {
                recipients.push(recipient.emailAddress.address.toLowerCase());
            }
        });
    } else {
        // Group email - return first recipient or group identifier
        if (allRecipients.length > 0 && allRecipients[0].emailAddress?.address) {
            recipients.push(allRecipients[0].emailAddress.address.toLowerCase());
        }
    }

    return recipients;
}

function extractSender(email) {
    const sender = email.from?.emailAddress?.address || 
                   email.sender?.emailAddress?.address;
    return sender ? sender.toLowerCase() : null;
}

function initContactStats() {
    return {
        sent: 0,
        received: 0,
        repliedTo: 0,
        ignored: 0,
        avgResponseTime: null,
        responseTimes: []
    };
}


function calculateResponseMetrics(contactStats, conversations, config) {
    // Reset metrics that will be recalculated
    Object.keys(contactStats).forEach(contact => {
        contactStats[contact].repliedTo = 0; // How many times they replied to you
        contactStats[contact].ignored = 0;   // How many times they ignored you
        contactStats[contact].responseTimes = [];
    });

    Object.values(conversations).forEach(thread => {
        if (thread.length < 1) return;

        thread.sort((a, b) => a.timestamp - b.timestamp);

        const sentMessageStatus = {}; // Key: message index, Value: boolean (replied or not)

        for (let i = 0; i < thread.length; i++) {
            // If it's an email you sent
            if (thread[i].type === 'sent') {
                const sentMessage = thread[i];
                const recipient = sentMessage.recipient;
                sentMessageStatus[i] = false; // Assume ignored until a reply is found

                // Look for a reply in the rest of the thread
                for (let j = i + 1; j < thread.length; j++) {
                    const potentialReply = thread[j];
                    // A reply is a received message from the person you sent the email to
                    if (potentialReply.type === 'received' && potentialReply.sender === recipient) {
                        sentMessageStatus[i] = true; // Mark as replied

                        if (contactStats[recipient]) {
                            const responseTime = calculateTimeDifference(
                                sentMessage.timestamp,
                                potentialReply.timestamp,
                                config.useBusinessHours
                            );
                            contactStats[recipient].responseTimes.push(responseTime);
                        }
                        // Stop looking for a reply for this specific sent message
                        break;
                    }
                }
            }
        }

        // Aggregate the results
        Object.keys(sentMessageStatus).forEach(index => {
            const messageIndex = parseInt(index);
            const sentMessage = thread[messageIndex];
            const recipient = sentMessage.recipient;

            if (contactStats[recipient]) {
                if (sentMessageStatus[messageIndex]) {
                    contactStats[recipient].repliedTo++;
                } else {
                    contactStats[recipient].ignored++;
                }
            }
        });
    });

    // Calculate final average response times
    Object.keys(contactStats).forEach(contact => {
        const stats = contactStats[contact];
        if (stats.responseTimes.length > 0) {
            const avgMinutes = stats.responseTimes.reduce((a, b) => a + b, 0) / stats.responseTimes.length;
            stats.avgResponseTime = formatResponseTime(avgMinutes);
        } else {
            stats.avgResponseTime = null;
        }
        delete stats.responseTimes; // Clean up temporary data
    });
}


function calculateTimeDifference(start, end, useBusinessHours) {
    let diffMs = end - start;
    let diffMinutes = diffMs / (1000 * 60);

    if (!useBusinessHours) {
        return diffMinutes;
    }

    // Calculate business hours only (Monday-Friday, 9 AM - 5 PM)
    let businessMinutes = 0;
    let current = new Date(start);
    const endTime = new Date(end);

    while (current < endTime) {
        const day = current.getDay();
        const hour = current.getHours();

        // Check if it's a weekday and within business hours
        if (day >= 1 && day <= 5 && hour >= 9 && hour < 17) {
            businessMinutes++;
        }

        current.setMinutes(current.getMinutes() + 1);

        // Optimization: skip non-business hours in larger chunks
        if (day === 0 || day === 6 || hour < 9 || hour >= 17) {
            if (day === 0 || day === 6) {
                // Skip to Monday 9 AM
                current.setDate(current.getDate() + (day === 0 ? 1 : 2));
                current.setHours(9, 0, 0, 0);
            } else if (hour >= 17) {
                // Skip to next day 9 AM
                current.setDate(current.getDate() + 1);
                current.setHours(9, 0, 0, 0);
            } else {
                // Skip to 9 AM same day
                current.setHours(9, 0, 0, 0);
            }
        }
    }

    return businessMinutes;
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

function useMockData() {
    // Fallback mock data for testing
    const mockStats = {
        "john.doe@company.com": { 
            sent: 45, 
            received: 52, 
            repliedTo: 25, 
            ignored: 20, 
            avgResponseTime: "8.5h" 
        },
        "jane.smith@company.com": { 
            sent: 28, 
            received: 35, 
            repliedTo: 25, 
            ignored: 3, 
            avgResponseTime: "45m" 
        }
    };

    chrome.storage.local.set({ emailStats: mockStats }, () => {
        console.log('Mock email data loaded for development.');
    });
}