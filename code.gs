// --- GLOBAL CONFIGURATION ---
const CONFIG = {
  KEYWORDS: {
    STRATEGIC_PLANNING: ['strategy', 'strategic', 'planning', 'roadmap', 'qbr', 'okr', 'goals', 'vision', 'budget', 'forecast', 'initiative'],
    LEADERSHIP_COMMS: ['1:1', 'one-on-one', 'team sync', 'all-hands', 'mentoring', 'coaching', 'feedback', 'review', 'hiring', 'headcount', 'development', 'update', 'internal comms'],
    OPERATIONAL_EXECUTION: ['support', 'ticket', 'bug', 'fix', 'deploy', 'release', 'ops', 'maintenance', 'troubleshoot', 'incident', 'report', 'standup', 'daily sync'],
    CLIENT_STAKEHOLDER: ['client:', 'customer:', 'partner:', 'q&a', 'demo', 'proposal', 'contract', 'negotiation', 'feedback from', 'follow-up with']
  },
  SCORING: {
    BASE_SCORE: 10,
    KEYWORD_HIT_VALUE: 15,
    COLLABORATION_BONUS_MEETING: 20,
    COLLABORATION_BONUS_EMAIL: 10,
    EXTERNAL_COMM_BONUS: 15,
    MAX_SCORE: 100
  },
  STYLE: {
    HEADER: {
      [DocumentApp.Attribute.FONT_FAMILY]: 'Calibri',
      [DocumentApp.Attribute.FONT_SIZE]: 10,
      [DocumentApp.Attribute.ITALIC]: true
    },
    TITLE: {
      [DocumentApp.Attribute.FONT_FAMILY]: 'Calibri',
      [DocumentApp.Attribute.FONT_SIZE]: 24,
      [DocumentApp.Attribute.BOLD]: true
    },
    H1: {
      [DocumentApp.Attribute.FONT_FAMILY]: 'Calibri',
      [DocumentApp.Attribute.FONT_SIZE]: 16,
      [DocumentApp.Attribute.BOLD]: true,
      [DocumentApp.Attribute.FOREGROUND_COLOR]: '#00558C'
    },
    H2: {
      [DocumentApp.Attribute.FONT_FAMILY]: 'Calibri',
      [DocumentApp.Attribute.FONT_SIZE]: 12,
      [DocumentApp.Attribute.BOLD]: true
    },
    TABLE_HEADER_CELL: {
      [DocumentApp.Attribute.BOLD]: true,
      [DocumentApp.Attribute.BACKGROUND_COLOR]: '#F3F3F3'
    },
    GREEN_TEXT: { [DocumentApp.Attribute.FOREGROUND_COLOR]: '#38761D' },
    ORANGE_TEXT: { [DocumentApp.Attribute.FOREGROUND_COLOR]: '#B45F06' },
    RED_TEXT: { [DocumentApp.Attribute.FOREGROUND_COLOR]: '#990000' }
  },
  ANALYSIS: {
    TOP_N_ACTIVITIES: 5,
    MAX_GMAIL_THREADS: 250,
    MAX_GMAIL_MESSAGES: 500,
    WORK_HOURS: { START: 8, END: 18 },
    MIN_DESCRIPTION_LENGTH: 10,
    BATCH_SIZE: 50
  },
  LIMITS: {
    MAX_DATE_RANGE_DAYS: 365,
    MIN_NAME_LENGTH: 2,
    MAX_TOPIC_LENGTH: 100
  }
};

const ROLE_PRESETS = {
  "EXECUTIVE": "50,35,5,10",
  "MANAGER": "20,40,25,15",
  "INDIVIDUAL_CONTRIBUTOR": "10,15,60,15"
};

// === UTILITY FUNCTIONS ===
/**
 * Validates email format
 */
function isValidEmail(email) {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

/**
 * Validates date range
 */
function validateDateRange(startDate, endDate) {
  const now = new Date();
  const daysDiff = Math.ceil((endDate - startDate) / (1000 * 60 * 60 * 24));
  
  if (startDate >= endDate) {
    throw new Error('Start date must be before end date');
  }
  if (endDate > now) {
    throw new Error('End date cannot be in the future');
  }
  if (daysDiff > CONFIG.LIMITS.MAX_DATE_RANGE_DAYS) {
    throw new Error(`Date range cannot exceed ${CONFIG.LIMITS.MAX_DATE_RANGE_DAYS} days`);
  }
}

/**
 * Sanitizes text input
 */
function sanitizeText(text) {
  if (!text) return '';
  return text.toString().trim().replace(/[<>\"']/g, '');
}

/**
 * Safe string contains check
 */
function safeContains(text, searchTerm) {
  if (!text || !searchTerm) return false;
  return text.toLowerCase().includes(searchTerm.toLowerCase());
}

// === UI FUNCTIONS ===
/**
 * The trigger function that builds the initial UI when the add-on is opened in Gmail.
 */
function onGmailHomepage(e) {
  try {
    return createWelcomeCard();
  } catch (error) {
    console.error('Error in onGmailHomepage:', error);
    return createErrorCard('Failed to load add-on', error.message);
  }
}

/**
 * Creates the initial welcome card that explains the add-on's purpose.
 */
function createWelcomeCard() {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Work Summary Generator'))
    .addSection(
      CardService.newCardSection()
        .addWidget(CardService.newTextParagraph().setText(
          'This tool analyzes your Gmail and Calendar activity to create a professional, VP-ready work summary, perfect for performance reviews.'))
        .addWidget(CardService.newButtonSet().addButton(
          CardService.newTextButton()
            .setText('Create New Report')
            .setOnClickAction(CardService.newAction().setFunctionName('navigateToConfigurationCard'))
            .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        ))
    ).build();
}

/**
 * Action function to navigate from the welcome screen to the main form.
 */
function navigateToConfigurationCard() {
  try {
    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().pushCard(createConfigurationCard()))
      .build();
  } catch (error) {
    console.error('Error navigating to configuration:', error);
    return createErrorResponse('Navigation failed', error.message);
  }
}

/**
 * Creates the main UI card with all the user-friendly input fields.
 */
function createConfigurationCard(e = {}) {
  try {
    const formInputs = e.formInputs || {};
    const selectedRole = formInputs.role_select ? formInputs.role_select[0] : '';
    const weights = selectedRole ? ROLE_PRESETS[selectedRole] : (formInputs.weights ? formInputs.weights[0] : '');
    const currentUserEmail = Session.getActiveUser().getEmail();

    const card = CardService.newCardBuilder();
    card.setHeader(CardService.newCardHeader().setTitle('Create New Report').setSubtitle('Step 1: Configure Your Report'));

    // --- Section 1: Who and When ---
    const whoWhenSection = CardService.newCardSection().setHeader('Who & When');
    
    whoWhenSection.addWidget(
      CardService.newTextInput()
        .setFieldName('fullName')
        .setTitle("Person's Full Name")
        .setValue(formInputs.fullName ? formInputs.fullName[0] : '')
    );
    
    whoWhenSection.addWidget(
      CardService.newTextInput()
        .setFieldName('userEmail')
        .setTitle('Email to Analyze')
        .setValue(formInputs.userEmail ? formInputs.userEmail[0] : currentUserEmail)
    );
    
    whoWhenSection.addWidget(
      CardService.newTextParagraph()
        .setText("<i><b>Note:</b> For security, this script can only analyze the data of the person running it.</i>")
    );
    
    whoWhenSection.addWidget(
      CardService.newDatePicker()
        .setFieldName('startDate')
        .setTitle('Start Date')
    );
    
    whoWhenSection.addWidget(
      CardService.newDatePicker()
        .setFieldName('endDate')
        .setTitle('End Date')
    );
    
    card.addSection(whoWhenSection);

    // --- Section 2 for Topical Summary ---
    const topicSection = CardService.newCardSection().setHeader('Topical Deep Dive (Optional)');
    
    topicSection.addWidget(
      CardService.newTextInput()
        .setFieldName('topic')
        .setTitle('Summarize a topic')
        .setHint("e.g., 'Project Phoenix' or 'Q4 Budget'")
        .setValue(formInputs.topic ? formInputs.topic[0] : '')
    );
        
    topicSection.addWidget(
      CardService.newTextInput()
        .setFieldName('personEmail')
        .setTitle('Filter topic by person (Optional)')
        .setHint("e.g., jane.doe@example.com")
        .setValue(formInputs.personEmail ? formInputs.personEmail[0] : '')
    );

    card.addSection(topicSection);

    // --- Section 3: Strategy Weights ---
    const weightsSection = CardService.newCardSection().setHeader('Define Role Focus (Strategy Weights)');
    
    weightsSection.addWidget(
      CardService.newTextParagraph()
        .setText("This tells the tool what's most important for the role. Think of it like a <b>budget for professional focus.</b> Start with a preset and adjust if needed.")
    );

    const roleSelect = CardService.newSelectionInput()
      .setFieldName('role_select')
      .setTitle("Select a Role Profile (to auto-fill weights)")
      .setType(CardService.SelectionInputType.DROPDOWN)
      .addItem("Select a preset...", "", selectedRole === "")
      .addItem("Executive / Director", "EXECUTIVE", selectedRole === "EXECUTIVE")
      .addItem("Manager / Team Lead", "MANAGER", selectedRole === "MANAGER")
      .addItem("Individual Contributor", "INDIVIDUAL_CONTRIBUTOR", selectedRole === "INDIVIDUAL_CONTRIBUTOR")
      .setOnChangeAction(CardService.newAction().setFunctionName('handleRoleChange'));
    
    weightsSection.addWidget(roleSelect);
    
    weightsSection.addWidget(
      CardService.newTextInput()
        .setFieldName('weights')
        .setTitle('Strategy Weights (Strategic, Leadership, Operational, Client)')
        .setValue(weights)
        .setHint('Must be 4 numbers separated by commas, totaling 100')
    );

    card.addSection(weightsSection);

    // --- Section 4: Generate Button ---
    card.setFixedFooter(
      CardService.newFixedFooter().setPrimaryButton(
        CardService.newTextButton()
          .setText('Generate Report')
          .setOnClickAction(CardService.newAction().setFunctionName('createReportAction'))
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      )
    );

    return card.build();
  } catch (error) {
    console.error('Error creating configuration card:', error);
    return createErrorCard('Configuration Error', error.message);
  }
}

/**
 * Action handler that re-renders the configuration card when a role preset is chosen.
 */
function handleRoleChange(e) {
  try {
    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().updateCard(createConfigurationCard(e)))
      .build();
  } catch (error) {
    console.error('Error handling role change:', error);
    return createErrorResponse('Role Change Failed', error.message);
  }
}

/**
 * The action handler function that is called when the "Generate Report" button is clicked.
 */
function createReportAction(e) {
  try {
    const formInputs = e.formInputs;
    
    // Validate required fields
    if (!formInputs.fullName || !formInputs.fullName[0]) {
      return createErrorResponse('Validation Error', 'Full name is required');
    }
    if (!formInputs.startDate || !formInputs.endDate) {
      return createErrorResponse('Validation Error', 'Both start and end dates are required');
    }
    if (!formInputs.weights || !formInputs.weights[0]) {
      return createErrorResponse('Validation Error', 'Strategy weights are required');
    }

    const userEmail = sanitizeText(formInputs.userEmail[0]);
    const fullName = sanitizeText(formInputs.fullName[0]);
    const startDateMs = formInputs.startDate[0].msSinceEpoch;
    const endDateMs = formInputs.endDate[0].msSinceEpoch;
    const weightsStr = sanitizeText(formInputs.weights[0]);
    const topic = formInputs.topic ? sanitizeText(formInputs.topic[0]) : null;
    const personEmail = formInputs.personEmail ? sanitizeText(formInputs.personEmail[0]) : null;

    // Validate inputs
    if (fullName.length < CONFIG.LIMITS.MIN_NAME_LENGTH) {
      return createErrorResponse('Validation Error', 'Name must be at least 2 characters');
    }
    
    if (!isValidEmail(userEmail)) {
      return createErrorResponse('Validation Error', 'Invalid email format');
    }
    
    if (personEmail && !isValidEmail(personEmail)) {
      return createErrorResponse('Validation Error', 'Invalid person email format');
    }
    
    if (topic && topic.length > CONFIG.LIMITS.MAX_TOPIC_LENGTH) {
      return createErrorResponse('Validation Error', `Topic must be less than ${CONFIG.LIMITS.MAX_TOPIC_LENGTH} characters`);
    }

    const startDate = new Date(startDateMs);
    const endDate = new Date(endDateMs);
    
    validateDateRange(startDate, endDate);
    
    const userInputs = {
      userEmail: userEmail,
      fullName: fullName,
      startDate: startDate,
      endDate: endDate,
      userDomain: userEmail.split('@')[1],
      weights: {},
      topic: topic,
      personEmail: personEmail
    };

    // Parse and validate weights
    const weightsArr = weightsStr.split(',').map(w => {
      const parsed = parseInt(w.trim(), 10);
      if (isNaN(parsed) || parsed < 0 || parsed > 100) {
        throw new Error('Invalid weight value: ' + w.trim());
      }
      return parsed;
    });
    
    if (weightsArr.length !== 4) {
      return createErrorResponse('Validation Error', 'Must provide exactly 4 weight values');
    }
    
    const weightSum = weightsArr.reduce((a, b) => a + b, 0);
    if (weightSum !== 100) {
      return createErrorResponse('Validation Error', `Weights must sum to 100, got ${weightSum}`);
    }
    
    userInputs.weights = {
      STRATEGIC_PLANNING: weightsArr[0],
      LEADERSHIP_COMMS: weightsArr[1],
      OPERATIONAL_EXECUTION: weightsArr[2],
      CLIENT_STAKEHOLDER: weightsArr[3]
    };

    // Generate report
    const rawData = fetchData(userInputs.startDate, userInputs.endDate, userInputs.userEmail);
    const analysis = analyzeData(rawData, userInputs.weights, userInputs.userDomain, userInputs.topic, userInputs.personEmail);
    const docUrl = generateReport(analysis, userInputs);

    const successCard = createSuccessCard(docUrl, analysis.metrics);
    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().pushCard(successCard))
      .build();

  } catch (error) {
    console.error('Error in createReportAction:', error);
    return createErrorResponse('Report Generation Failed', error.message);
  }
}

/**
 * Creates the final success card with a link to the report.
 */
function createSuccessCard(docUrl, metrics) {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Report Generation Complete!'))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newKeyValue()
        .setTopLabel('Status')
        .setContent('✅ Success')
        .setIcon(CardService.Icon.CONFIRMATION))
      .addWidget(CardService.newKeyValue()
        .setTopLabel('Strategic Focus Score')
        .setContent(`${metrics.weightedStrategicFocus.toFixed(1)}%`))
      .addWidget(CardService.newKeyValue()
        .setTopLabel('Activities Analyzed')
        .setContent(metrics.totalActivities.toString()))
    )
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText('Open Report')
          .setOpenLink(CardService.newOpenLink().setUrl(docUrl))
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED))
        .addButton(CardService.newTextButton()
          .setText('Create Another')
          .setOnClickAction(CardService.newAction().setFunctionName('handleCreateAnother')))
      )
    ).build();
}

/**
 * Action handler to go back to the welcome screen.
 */
function handleCreateAnother() {
  try {
    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().popToRoot())
      .build();
  } catch (error) {
    console.error('Error in handleCreateAnother:', error);
    return createErrorResponse('Navigation Error', error.message);
  }
}

/**
 * Creates an error card for display
 */
function createErrorCard(title, message) {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle(title).setSubtitle('Please try again'))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph().setText(message))
      .addWidget(CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText('Try Again')
          .setOnClickAction(CardService.newAction().setFunctionName('navigateToConfigurationCard')))
      )
    ).build();
}

/**
 * Creates an error response for actions
 */
function createErrorResponse(title, message) {
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText(`${title}: ${message}`))
    .build();
}

// === CORE LOGIC FUNCTIONS ===

/**
 * Fetches calendar and email data for the specified date range and user.
 * Fixed the main bug: replaced non-existent isOwnedBy() method
 */
function fetchData(startDate, endDate, userEmail) {
  const currentUserEmail = Session.getActiveUser().getEmail();
  if (userEmail !== currentUserEmail) {
    throw new Error(`Security restriction: This script can only analyze data for ${currentUserEmail}`);
  }

  let activities = [];
  
  try {
    // Fetch calendar events - FIXED: removed isOwnedBy() call
    const calendar = CalendarApp.getDefaultCalendar();
    const events = calendar.getEvents(startDate, endDate);
    
    console.log(`Found ${events.length} calendar events`);
    
    for (const event of events) {
      try {
        const userGuest = event.getGuestByEmail(userEmail);
        const userStatus = userGuest ? userGuest.getGuestStatus() : null;
        
        // Include event if user accepted OR if no guest status (meaning they're the owner)
        if (userStatus === CalendarApp.GuestStatus.YES || userStatus === null) {
          const attendeeEmails = [];
          try {
            const guestList = event.getGuestList();
            for (const guest of guestList) {
              const guestEmail = guest.getEmail();
              if (guestEmail) {
                attendeeEmails.push(guestEmail);
              }
            }
          } catch (e) {
            console.warn('Error getting guest list for event:', event.getTitle(), e);
          }

          activities.push({
            type: 'Meeting',
            id: event.getId(),
            date: event.getStartTime(),
            title: sanitizeText(event.getTitle()) || 'Untitled Event',
            description: sanitizeText(event.getDescription()) || '',
            attendees: attendeeEmails
          });
        }
      } catch (e) {
        console.warn('Error processing calendar event:', e);
      }
    }
    
    // Fetch Gmail threads with better error handling
    try {
      const gmailQuery = `from:${userEmail} after:${formatDateForGmail(startDate)} before:${formatDateForGmail(endDate)}`;
      console.log('Gmail query:', gmailQuery);
      
      const threads = GmailApp.search(gmailQuery, 0, CONFIG.ANALYSIS.MAX_GMAIL_THREADS);
      console.log(`Found ${threads.length} email threads`);
      
      for (const thread of threads) {
        try {
          const messages = thread.getMessages();
          
          // Process messages in batches to avoid timeout
          for (let i = 0; i < Math.min(messages.length, CONFIG.ANALYSIS.MAX_GMAIL_MESSAGES); i++) {
            const message = messages[i];
            
            try {
              const messageDate = message.getDate();
              if (messageDate >= startDate && messageDate <= endDate) {
                const recipients = extractRecipients(message);
                
                activities.push({
                  type: 'Email',
                  id: message.getId(),
                  date: messageDate,
                  title: sanitizeText(message.getSubject()) || 'No Subject',
                  description: truncateText(sanitizeText(message.getPlainBody()), 500),
                  attendees: recipients
                });
              }
            } catch (e) {
              console.warn('Error processing message in thread:', e);
            }
          }
        } catch (e) {
          console.warn('Error processing email thread:', e);
        }
      }
    } catch (e) {
      console.error('Error fetching Gmail data:', e);
      // Don't throw here, continue with calendar data only
    }
    
  } catch (error) {
    console.error('Error in fetchData:', error);
    throw new Error(`Failed to fetch data: ${error.message}`);
  }
  
  console.log(`Total activities found: ${activities.length}`);
  return activities;
}

/**
 * Helper function to format date for Gmail search
 */
function formatDateForGmail(date) {
  const year = date.getFullYear();
  const month = (date.getMonth() + 1).toString().padStart(2, '0');
  const day = date.getDate().toString().padStart(2, '0');
  return `${year}/${month}/${day}`;
}

/**
 * Helper function to extract recipients from email message
 */
function extractRecipients(message) {
  const recipients = [];
  try {
    const to = message.getTo();
    const cc = message.getCc();
    const bcc = message.getBcc();
    
    [to, cc, bcc].forEach(field => {
      if (field) {
        field.split(',').forEach(email => {
          const trimmedEmail = email.trim();
          if (trimmedEmail && isValidEmail(trimmedEmail)) {
            recipients.push(trimmedEmail);
          }
        });
      }
    });
  } catch (e) {
    console.warn('Error extracting recipients:', e);
  }
  return recipients;
}

/**
 * Helper function to truncate text
 */
function truncateText(text, maxLength) {
  if (!text || text.length <= maxLength) return text;
  return text.substring(0, maxLength) + '...';
}

/**
 * Main analysis function with improved error handling and logic
 */
function analyzeData(activities, weights, userDomain, topic, personEmail) {
  const analysis = {
    activities: [],
    metrics: {},
    distribution: {
      STRATEGIC_PLANNING: { count: 0, totalScore: 0, items: [] },
      LEADERSHIP_COMMS: { count: 0, totalScore: 0, items: [] },
      OPERATIONAL_EXECUTION: { count: 0, totalScore: 0, items: [] },
      CLIENT_STAKEHOLDER: { count: 0, totalScore: 0, items: [] },
      UNCATEGORIZED: { count: 0, totalScore: 0, items: [] }
    },
    workPatterns: {
      byHour: Array(24).fill(0),
      commSplit: { internal: 0, external: 0 }
    },
    recommendations: [],
    topicSummary: null
  };

  if (!activities || activities.length === 0) {
    analysis.metrics = {
      totalActivities: 0,
      totalMeetings: 0,
      totalEmails: 0,
      averageRelevance: 0,
      weightedStrategicFocus: 0
    };
    analysis.recommendations.push("No activities found for the selected period.");
    return analysis;
  }

  // Run topic summary if a topic is provided
  if (topic && topic.length > 0) {
    analysis.topicSummary = summarizeTopicFromActivities(topic, personEmail, activities);
  }

  // Perform the main work categorization and scoring
  for (const activity of activities) {
    try {
      const textCorpus = `${activity.title.toLowerCase()} ${activity.description.toLowerCase()}`;
      let bestCategory = 'UNCATEGORIZED';
      let maxHits = 0;

      // Find the category with the most keyword hits
      for (const category in CONFIG.KEYWORDS) {
        const keywords = CONFIG.KEYWORDS[category];
        const hits = keywords.filter(keyword => safeContains(textCorpus, keyword)).length;
        if (hits > maxHits) {
          maxHits = hits;
          bestCategory = category;
        }
      }

      activity.category = bestCategory;

      // Calculate relevance score
      let score = CONFIG.SCORING.BASE_SCORE + (maxHits * CONFIG.SCORING.KEYWORD_HIT_VALUE);

      // Add collaboration bonuses
      const attendeeCount = activity.attendees ? activity.attendees.length : 0;
      if (activity.type === 'Meeting' && attendeeCount > 1) {
        score += CONFIG.SCORING.COLLABORATION_BONUS_MEETING;
      }
      if (activity.type === 'Email' && attendeeCount > 1) {
        score += CONFIG.SCORING.COLLABORATION_BONUS_EMAIL;
      }

      // Add external communication bonus
      const isExternal = activity.attendees && activity.attendees.some(email => {
        return email && email.includes('@') && !email.endsWith('@' + userDomain);
      });
      if (isExternal) {
        score += CONFIG.SCORING.EXTERNAL_COMM_BONUS;
      }

      activity.relevanceScore = Math.min(CONFIG.SCORING.MAX_SCORE, Math.round(score));
      analysis.activities.push(activity);

      // Update distribution
      const dist = analysis.distribution[bestCategory];
      dist.count++;
      dist.totalScore += activity.relevanceScore;
      dist.items.push(activity);

      // Update work patterns
      const hour = activity.date.getHours();
      if (hour >= 0 && hour < 24) {
        analysis.workPatterns.byHour[hour]++;
      }

      if (isExternal) {
        analysis.workPatterns.commSplit.external++;
      } else {
        analysis.workPatterns.commSplit.internal++;
      }
    } catch (e) {
      console.warn('Error analyzing activity:', activity.title, e);
    }
  }

  // Calculate metrics
  const totalActivities = activities.length;
  analysis.metrics.totalActivities = totalActivities;
  analysis.metrics.totalMeetings = activities.filter(a => a.type === 'Meeting').length;
  analysis.metrics.totalEmails = activities.filter(a => a.type === 'Email').length;
  analysis.metrics.averageRelevance = totalActivities > 0 ? 
    analysis.activities.reduce((sum, a) => sum + a.relevanceScore, 0) / totalActivities : 0;

  // Calculate weighted strategic focus
  let weightedScoreSum = 0;
  const totalWeight = Object.values(weights).reduce((s, w) => s + w, 0) || 100;

  for (const category in analysis.distribution) {
    if (weights[category]) {
      const categoryPercentage = totalActivities > 0 ? 
        (analysis.distribution[category].count / totalActivities) * 100 : 0;
      weightedScoreSum += categoryPercentage * (weights[category] / totalWeight);
    }
  }
  analysis.metrics.weightedStrategicFocus = weightedScoreSum;

  // Generate recommendations
  generateRecommendations(analysis, totalActivities);

  return analysis;
}

/**
 * Generates recommendations based on analysis
 */
function generateRecommendations(analysis, totalActivities) {
  if (totalActivities === 0) {
    analysis.recommendations.push("No activities found to analyze.");
    return;
  }

  const strategicPct = (analysis.distribution.STRATEGIC_PLANNING.count / totalActivities) * 100;
  const operationalPct = (analysis.distribution.OPERATIONAL_EXECUTION.count / totalActivities) * 100;
  const leadershipPct = (analysis.distribution.LEADERSHIP_COMMS.count / totalActivities) * 100;

  if (strategicPct > 40) {
    analysis.recommendations.push("✅ Strong focus on strategic initiatives is evident.");
  }
  
  if (operationalPct > 50) {
    analysis.recommendations.push("📈 Significant time on operational tasks. Consider delegation opportunities.");
  }
  
  if (leadershipPct > 30) {
    analysis.recommendations.push("👥 Good investment in leadership and team development.");
  }
  
  const outOfHoursActivities = analysis.workPatterns.byHour.slice(0, CONFIG.ANALYSIS.WORK_HOURS.START)
    .concat(analysis.workPatterns.byHour.slice(CONFIG.ANALYSIS.WORK_HOURS.END))
    .reduce((sum, count) => sum + count, 0);
  
  if (outOfHoursActivities > totalActivities * 0.2) {
    analysis.recommendations.push("🕐 High after-hours activity detected. Consider work-life balance.");
  }
  
  if (analysis.recommendations.length === 0) {
    analysis.recommendations.push("📊 Work distribution appears balanced across categories.");
  }
}

/**
 * Finds and summarizes emails related to a topic, with an optional filter for a specific person.
 */
function summarizeTopicFromActivities(topic, personEmail, activities) {
  const lowerCaseTopic = topic.toLowerCase();
  
  // First, filter by the topic
  let relevantEmails = activities.filter(activity => 
    activity.type === 'Email' &&
    (safeContains(activity.title, lowerCaseTopic) || 
     safeContains(activity.description, lowerCaseTopic))
  );

  // If a person's email is provided, filter the results further
  if (personEmail && personEmail.length > 0) {
    const lowerCasePersonEmail = personEmail.toLowerCase();
    relevantEmails = relevantEmails.filter(email => 
      email.attendees.some(attendee => attendee.toLowerCase().includes(lowerCasePersonEmail))
    );
  }

  if (relevantEmails.length === 0) {
    return null;
  }

  const uniqueParticipants = new Set();
  relevantEmails.forEach(email => {
    email.attendees.forEach(p => {
      if (p) uniqueParticipants.add(p);
    });
  });

  return {
    count: relevantEmails.length,
    topSubjects: relevantEmails.slice(0, 5).map(email => email.title),
    participants: Array.from(uniqueParticipants)
  };
}

/**
 * Generates the report document with improved error handling
 */
function generateReport(analysis, inputs) {
  try {
    const docName = `Work Summary - ${inputs.fullName} (${inputs.startDate.toISOString().slice(0, 10)})`;
    const doc = DocumentApp.create(docName);
    const body = doc.getBody();
    
    // Clear default content
    body.clear();
    
    // Title
    body.appendParagraph(doc.getName()).setAttributes(CONFIG.STYLE.TITLE);
    body.appendParagraph(`Analysis Period: ${inputs.startDate.toLocaleDateString()} to ${inputs.endDate.toLocaleDateString()}`)
        .setAttributes(CONFIG.STYLE.HEADER);
    body.appendParagraph(''); // Empty line
    
    // Insert the topic summary section if it exists
    if (analysis.topicSummary) {
      insertTopicSummary(body, analysis.topicSummary, inputs.topic, inputs.personEmail);
    }
    
    // Executive Summary
    insertExecutiveSummary(body, analysis.metrics);
    
    // Performance Metrics
    body.appendParagraph('Performance Metrics').setAttributes(CONFIG.STYLE.H1);
    insertMetricsTable(body, analysis.metrics);
    
    // Activity Distribution
    body.appendParagraph('Activity Distribution').setAttributes(CONFIG.STYLE.H1);
    insertDistributionTable(body, analysis.distribution, analysis.metrics.totalActivities);
    
    // Detailed Activities
    body.appendParagraph('Detailed Activities').setAttributes(CONFIG.STYLE.H1);
    insertDetailedActivities(body, analysis.distribution);
    
    // Work Patterns
    body.appendParagraph('Work Patterns').setAttributes(CONFIG.STYLE.H1);
    insertWorkPatterns(body, analysis.workPatterns, analysis.metrics.totalActivities);
    
    // Executive Recommendations
    body.appendParagraph('Executive Recommendations').setAttributes(CONFIG.STYLE.H1);
    analysis.recommendations.forEach(rec => {
      body.appendListItem(rec);
    });
    
    doc.saveAndClose();
    return doc.getUrl();
  } catch (error) {
    console.error('Error generating report:', error);
    throw new Error(`Failed to generate report: ${error.message}`);
  }
}

/**
 * Writes the topic summary to the Google Doc, now with a dynamic title.
 */
function insertTopicSummary(body, topicSummary, topic, personEmail) {
  let heading = 'Topical Deep Dive: "' + topic + '"';
  if (personEmail && personEmail.length > 0) {
    heading += ' with ' + personEmail;
  }
  
  body.appendParagraph(heading).setAttributes(CONFIG.STYLE.H1);
  
  const tableData = [
    ['Metric', 'Details'],
    ['Matching Email Threads', topicSummary.count.toString()],
    ['Key Conversation Subjects', topicSummary.topSubjects.join('\n') || 'N/A'],
    ['Key Participants', topicSummary.participants.slice(0, 10).join(', ') + 
                        (topicSummary.participants.length > 10 ? '...' : '')]
  ];
  createStyledTable(body, tableData);
}

/**
 * Inserts executive summary with key metrics
 */
function insertExecutiveSummary(body, metrics) {
  const data = [
    ['Strategic Focus Score', `${metrics.weightedStrategicFocus.toFixed(1)}%`, 'Alignment with role priorities.'],
    ['Avg. Relevance Score', `${metrics.averageRelevance.toFixed(1)} / 100`, 'Average impact of activities.'],
    ['Total Activities', `${metrics.totalActivities}`, `Meetings & sent emails analyzed.`]
  ];
  
  const table = body.appendTable(data);
  table.setBorderWidth(0);
  
  for (let i = 0; i < data.length; i++) {
    table.getCell(i, 0).setAttributes(CONFIG.STYLE.H2);
    table.getCell(i, 1).getChild(0).asParagraph()
         .setAttributes(CONFIG.STYLE.GREEN_TEXT)
         .setBold(true)
         .setFontSize(14);
    table.getCell(i, 2).getChild(0).asParagraph().setItalic(true);
  }
  body.appendParagraph('');
}

/**
 * Inserts metrics table
 */
function insertMetricsTable(body, metrics) {
  const data = [
    ['Metric', 'Value'],
    ['Total Activities', metrics.totalActivities.toString()],
    ['Meetings', metrics.totalMeetings.toString()],
    ['Emails Sent', metrics.totalEmails.toString()],
    ['Avg. Relevance', `${metrics.averageRelevance.toFixed(1)}%`]
  ];
  createStyledTable(body, data);
}

/**
 * Inserts distribution table
 */
function insertDistributionTable(body, distribution, total) {
  if (total === 0) {
    body.appendParagraph("No activities found to distribute.");
    return;
  }
  
  const data = [['Category', '% of Activities', 'Avg. Relevance']];
  const categoryMap = {
    STRATEGIC_PLANNING: '🗺️ Strategic',
    LEADERSHIP_COMMS: '👥 Leadership',
    OPERATIONAL_EXECUTION: '⚙️ Operational',
    CLIENT_STAKEHOLDER: '🤝 Client-Facing',
    UNCATEGORIZED: '📥 Uncategorized'
  };
  
  for (const cat in categoryMap) {
    const item = distribution[cat];
    if (item.count > 0) {
      const percentage = ((item.count / total) * 100).toFixed(1);
      const avgRelevance = (item.totalScore / item.count).toFixed(1);
      data.push([categoryMap[cat], `${percentage}%`, `${avgRelevance}%`]);
    }
  }
  
  createStyledTable(body, data);
}

/**
 * Inserts detailed activities by category
 */
function insertDetailedActivities(body, distribution) {
  const categoryMap = {
    STRATEGIC_PLANNING: '🗺️ Top Strategic Activities',
    LEADERSHIP_COMMS: '👥 Top Leadership Activities',
    OPERATIONAL_EXECUTION: '⚙️ Top Operational Activities',
    CLIENT_STAKEHOLDER: '🤝 Top Client-Facing Activities'
  };
  
  let hasActivities = false;
  
  for (const cat in categoryMap) {
    const items = distribution[cat].items;
    if (items.length > 0) {
      hasActivities = true;
      body.appendParagraph(categoryMap[cat]).setAttributes(CONFIG.STYLE.H2);
      
      const sorted = items.sort((a, b) => b.relevanceScore - a.relevanceScore)
                          .slice(0, CONFIG.ANALYSIS.TOP_N_ACTIVITIES);
      
      sorted.forEach(item => {
        const style = item.relevanceScore > 70 ? CONFIG.STYLE.GREEN_TEXT :
                     item.relevanceScore > 40 ? CONFIG.STYLE.ORANGE_TEXT :
                     CONFIG.STYLE.RED_TEXT;
        
        const listItem = body.appendListItem(`[${item.date.toLocaleDateString()}] ${item.title}`);
        listItem.appendText(` (Score: `)
                .appendText(`${item.relevanceScore}`).setAttributes(style).setBold(true)
                .appendText(`)`);
      });
    }
  }
  
  if (!hasActivities) {
    body.appendParagraph("No categorized activities to detail.");
  }
}

/**
 * Inserts work patterns analysis
 */
function insertWorkPatterns(body, patterns, total) {
  if (total === 0) {
    body.appendParagraph("No work patterns to analyze.");
    return;
  }
  
  const totalComm = patterns.commSplit.internal + patterns.commSplit.external;
  
  // Communication Focus
  body.appendParagraph('Communication Focus').setAttributes(CONFIG.STYLE.H2);
  if (totalComm > 0) {
    body.appendListItem(`Internal: ${((patterns.commSplit.internal / totalComm) * 100).toFixed(0)}%`);
    body.appendListItem(`External: ${((patterns.commSplit.external / totalComm) * 100).toFixed(0)}%`);
  } else {
    body.appendListItem("No communication data available.");
  }
  body.appendParagraph('');
  
  // Activity Timing
  let outOfHours = 0;
  patterns.byHour.forEach((count, hour) => {
    if (count > 0 && (hour < CONFIG.ANALYSIS.WORK_HOURS.START || hour >= CONFIG.ANALYSIS.WORK_HOURS.END)) {
      outOfHours += count;
    }
  });
  
  body.appendParagraph('Activity Timing').setAttributes(CONFIG.STYLE.H2);
  body.appendListItem(`Work outside standard hours: ${total > 0 ? ((outOfHours / total) * 100).toFixed(0) : 0}%`);
  
  const peakHour = patterns.byHour.indexOf(Math.max(...patterns.byHour));
  body.appendListItem(`Peak activity hour: ${peakHour}:00`);
}

/**
 * Creates a styled table with proper formatting
 */
function createStyledTable(body, data) {
  if (data.length < 2) {
    body.appendParagraph("No data available.");
    return;
  }
  
  const table = body.appendTable(data);
  
  // Style the header row
  const headerRow = table.getRow(0);
  for (let j = 0; j < headerRow.getNumCells(); j++) {
    headerRow.getCell(j).setAttributes(CONFIG.STYLE.TABLE_HEADER_CELL);
  }
  
  body.appendParagraph('');
}
