// ================================================================================
// WORK SUMMARY GENERATOR - CLEAN SYNTAX VERSION
// ================================================================================

// --- GLOBAL CONFIGURATION ---
var CONFIG = {
    KEYWORDS: {
        INNOVATION: {
            stems: ['innovat', 'creat', 'design', 'prototype', 'pilot', 'experiment', 'research', 'breakthrough', 'disrupt'],
            exact: ['mvp', 'r&d', 'poc', 'beta', 'alpha'],
            patterns: ['proof of concept', 'minimum viable', 'design thinking']
        },
        EXECUTION: {
            stems: ['deliver', 'ship', 'launch', 'complet', 'implement', 'deploy', 'achiev', 'finish', 'execut'],
            exact: ['go-live', 'rollout', 'milestone', 'deadline', 'kpi', 'roi', 'target'],
            patterns: ['project completion', 'goal completion', 'went live']
        },
        COLLABORATION: {
            stems: ['collaborat', 'partner', 'coordinat', 'align', 'sync', 'team', 'group', 'joint', 'shared'],
            exact: ['cross-functional', 'stakeholder', 'workshop', 'meeting', 'together'],
            patterns: ['working with', 'team effort', 'group effort']
        },
        LEADERSHIP: {
            stems: ['lead', 'mentor', 'coach', 'guid', 'direct', 'manag', 'supervis', 'influenc', 'inspir', 'motivat'],
            exact: ['1:1', 'one-on-one', 'feedback', 'delegation', 'vision', 'strategy'],
            patterns: ['decision making', 'strategic direction', 'team development']
        }
    },
    
    OUTCOME_KEYWORDS: {
        revenue: {
            stems: ['revenue', 'sales', 'profit', 'income', 'earning', 'monetiz', 'deal', 'contract'],
            exact: ['roi', 'margin', 'pricing', 'upsell', 'renewal'],
            multiplier: 2.5
        },
        customer_satisfaction: {
            stems: ['customer', 'client', 'user', 'satisfaction', 'experience', 'retention', 'success'],
            exact: ['nps', 'csat', 'churn', 'ux', 'ui'],
            multiplier: 2.0
        },
        efficiency: {
            stems: ['efficiency', 'productivity', 'automat', 'optim', 'streamlin', 'scale', 'speed'],
            exact: ['sla', 'turnaround', 'throughput'],
            multiplier: 1.8
        },
        market_position: {
            stems: ['market', 'competit', 'brand', 'growth', 'expansion', 'penetrat'],
            exact: ['market share', 'differentiation'],
            multiplier: 2.2
        }
    },
    
    SCORING: {
        BASE_SCORE: 20,
        STEM_MATCH_VALUE: 10,
        EXACT_MATCH_VALUE: 15,
        PATTERN_MATCH_VALUE: 20,
        COLLABORATION_BONUS: 15,
        OUTCOME_BASE_BONUS: 25,
        MAX_CATEGORY_SCORE: 100,
        MULTI_LABEL_THRESHOLD: 30
    },

    STYLE: {
        HEADER: { [DocumentApp.Attribute.FONT_FAMILY]: 'Calibri', [DocumentApp.Attribute.FONT_SIZE]: 10, [DocumentApp.Attribute.ITALIC]: true },
        TITLE: { [DocumentApp.Attribute.FONT_FAMILY]: 'Calibri', [DocumentApp.Attribute.FONT_SIZE]: 24, [DocumentApp.Attribute.BOLD]: true },
        H1: { [DocumentApp.Attribute.FONT_FAMILY]: 'Calibri', [DocumentApp.Attribute.FONT_SIZE]: 16, [DocumentApp.Attribute.BOLD]: true, [DocumentApp.Attribute.FOREGROUND_COLOR]: '#00558C' },
        H2: { [DocumentApp.Attribute.FONT_FAMILY]: 'Calibri', [DocumentApp.Attribute.FONT_SIZE]: 12, [DocumentApp.Attribute.BOLD]: true }
    },
    
    ANALYSIS: {
        WORK_HOURS: { START: 8, END: 18 }
    }
};

var ROLE_PRESETS = {
  "INDIVIDUAL_CONTRIBUTOR": { name: "Individual Contributor", weights: "25,45,25,5" },
  "MANAGER": { name: "Manager / Team Lead", weights: "25,25,30,20" },
  "EXECUTIVE": { name: "Executive / Director", weights: "25,20,20,35" }
};

// ================================================================================
// MAIN ENTRY POINTS
// ================================================================================

function onGmailHomepage(e) {
  try {
    var jobs = getActiveJobs();
    return createHomepageCard(jobs);
  } catch (error) {
    console.error('Error in onGmailHomepage:', error);
    return createErrorCard(error.toString());
  }
}

// ================================================================================
// JOB MANAGEMENT SYSTEM
// ================================================================================

function getActiveJobs() {
  var properties = PropertiesService.getUserProperties();
  var jobIdsString = properties.getProperty('activeJobIds');
  
  if (!jobIdsString) return [];
  
  var jobIds = JSON.parse(jobIdsString);
  var jobs = [];
  
  for (var i = 0; i < jobIds.length; i++) {
    var jobDataString = properties.getProperty(jobIds[i]);
    if (jobDataString) {
      var jobData = JSON.parse(jobDataString);
      jobs.push(jobData);
    }
  }
  
  return jobs;
}

function createJob(userInputs) {
  var properties = PropertiesService.getUserProperties();
  var jobId = 'job_' + new Date().getTime();
  
  var jobIdsString = properties.getProperty('activeJobIds');
  var jobIds = jobIdsString ? JSON.parse(jobIdsString) : [];
  jobIds.push(jobId);
  properties.setProperty('activeJobIds', JSON.stringify(jobIds));

  var jobData = {
    id: jobId,
    status: 'pending',
    userInputs: userInputs,
    startTime: new Date().toISOString(),
    progress: 'Job created, waiting to start...'
  };
  properties.setProperty(jobId, JSON.stringify(jobData));
  
  return jobId;
}

function updateJobStatus(jobId, status, progress, additionalData) {
  var properties = PropertiesService.getUserProperties();
  var jobDataString = properties.getProperty(jobId);
  
  if (jobDataString) {
    var jobData = JSON.parse(jobDataString);
    jobData.status = status;
    if (progress) jobData.progress = progress;
    if (additionalData) {
      for (var key in additionalData) {
        jobData[key] = additionalData[key];
      }
    }
    properties.setProperty(jobId, JSON.stringify(jobData));
  }
}

function removeJob(jobId) {
  var properties = PropertiesService.getUserProperties();
  
  var jobIdsString = properties.getProperty('activeJobIds');
  if (jobIdsString) {
    var jobIds = JSON.parse(jobIdsString);
    var filteredIds = [];
    for (var i = 0; i < jobIds.length; i++) {
      if (jobIds[i] !== jobId) {
        filteredIds.push(jobIds[i]);
      }
    }
    properties.setProperty('activeJobIds', JSON.stringify(filteredIds));
  }
  
  properties.deleteProperty(jobId);
}

// ================================================================================
// USER INTERFACE
// ================================================================================

function createHomepageCard(jobs) {
  var builder = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Work Summary Generator'));

  var newReportSection = CardService.newCardSection()
    .setHeader("Create New Report")
    .setCollapsible(true);
  
  newReportSection.addWidget(
    CardService.newTextButton()
      .setText('Start New Report')
      .setOnClickAction(CardService.newAction().setFunctionName('navigateToConfigurationCard'))
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
  );
  
  builder.addSection(newReportSection);

  if (jobs.length > 0) {
    var statusSection = CardService.newCardSection().setHeader("Current Jobs");
    
    for (var i = 0; i < jobs.length; i++) {
      var job = jobs[i];
      var statusInfo = getStatusInfo(job.status);
      var timeStr = new Date(job.startTime).toLocaleString();
      
      var jobWidget = CardService.newKeyValue()
        .setTopLabel(job.userInputs.fullName + ' - ' + timeStr)
        .setContent(statusInfo.icon + ' ' + statusInfo.text)
        .setMultiline(true);
        
      if (job.progress) {
        jobWidget.setBottomLabel(job.progress);
      }
        
      jobWidget.setButton(
        CardService.newTextButton()
          .setText("View Details")
          .setOnClickAction(
            CardService.newAction()
              .setFunctionName("viewJobDetails")
              .setParameters({jobId: job.id})
          )
      );
      
      statusSection.addWidget(jobWidget);
    }
    
    builder.addSection(statusSection);
  }

  return builder.build();
}

function getStatusInfo(status) {
  if (status === 'pending') return { icon: '⏳', text: 'Pending' };
  if (status === 'processing') return { icon: '🔄', text: 'Processing' };
  if (status === 'completed') return { icon: '✅', text: 'Completed' };
  if (status === 'viewed') return { icon: '👁️', text: 'Viewed' };
  if (status === 'failed') return { icon: '❌', text: 'Failed' };
  return { icon: '❓', text: 'Unknown' };
}
/**
 * Creates the configuration card for new reports
 */
function createConfigurationCard() {
  const currentUserEmail = Session.getActiveUser().getEmail();

  const builder = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Report Configuration'));

  // Personal Information Section
  const personalSection = CardService.newCardSection().setHeader('Personal Information');
  
  personalSection.addWidget(
    CardService.newTextInput()
      .setFieldName('fullName')
      .setTitle("Full Name")
      .setHint("Your name as it should appear on the report")
  );

  personalSection.addWidget(
    CardService.newTextInput()
      .setFieldName('userEmail')
      .setTitle('Email to Analyze')
      .setValue(currentUserEmail)
      .setHint("The email account to analyze for work activities")
  );

  personalSection.addWidget(
    CardService.newTextInput()
      .setFieldName('notificationEmail')
      .setTitle('Notification Email')
      .setValue(currentUserEmail)
      .setHint("Where to send the completed report link")
  );

  builder.addSection(personalSection);

  // Date Range Section
  const dateSection = CardService.newCardSection().setHeader('Analysis Period');
  
  dateSection.addWidget(
    CardService.newDatePicker()
      .setFieldName('startDate')
      .setTitle('Start Date')
  );

  dateSection.addWidget(
    CardService.newDatePicker()
      .setFieldName('endDate')
      .setTitle('End Date')
  );

  builder.addSection(dateSection);

  // Role Profile Section
  const roleSection = CardService.newCardSection().setHeader('Role Profile');

  const roleSelect = CardService.newSelectionInput()
    .setFieldName('role_select')
    .setTitle("Role Type")
    .setType(CardService.SelectionInputType.DROPDOWN)
    .addItem(ROLE_PRESETS.INDIVIDUAL_CONTRIBUTOR.name, "INDIVIDUAL_CONTRIBUTOR", true)
    .addItem(ROLE_PRESETS.MANAGER.name, "MANAGER", false)
    .addItem(ROLE_PRESETS.EXECUTIVE.name, "EXECUTIVE", false);
  
  roleSection.addWidget(roleSelect);
  roleSection.addWidget(
    CardService.newTextParagraph()
      .setText("This determines how your activities are weighted for Innovation, Execution, Collaboration, and Leadership.")
  );
  
  builder.addSection(roleSection);

  // Generate Button
  builder.setFixedFooter(
    CardService.newFixedFooter().setPrimaryButton(
      CardService.newTextButton()
        .setText('Generate Report')
        .setOnClickAction(CardService.newAction().setFunctionName('createReportAction'))
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    )
  );

  return builder.build();
}

/**
 * Creates job detail view card
 */
function createJobDetailCard(jobData) {
  const builder = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Job Details'));

  const detailSection = CardService.newCardSection();
  
  const statusInfo = getStatusInfo(jobData.status);
  const timeStr = new Date(jobData.startTime).toLocaleString();
  
  detailSection.addWidget(
    CardService.newKeyValue()
      .setTopLabel('Name')
      .setContent(jobData.userInputs.fullName)
  );
  
  detailSection.addWidget(
    CardService.newKeyValue()
      .setTopLabel('Status')
      .setContent(`${statusInfo.icon} ${statusInfo.text}`)
  );
  
  detailSection.addWidget(
    CardService.newKeyValue()
      .setTopLabel('Started')
      .setContent(timeStr)
  );
  
  if (jobData.progress) {
    detailSection.addWidget(
      CardService.newKeyValue()
        .setTopLabel('Progress')
        .setContent(jobData.progress)
        .setMultiline(true)
    );
  }
  
  if (jobData.reportUrl) {
    detailSection.addWidget(
      CardService.newTextButton()
        .setText('Open Report')
        .setOnClickAction(
          CardService.newAction()
            .setFunctionName('openReportAndMarkViewed')
            .setParameters({jobId: jobData.id, reportUrl: jobData.reportUrl})
        )
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    );
  }
  
  if (jobData.errorMessage) {
    detailSection.addWidget(
      CardService.newKeyValue()
        .setTopLabel('Error Details')
        .setContent(jobData.errorMessage)
        .setMultiline(true)
    );
  }
  
  // Action buttons
  const actionSection = CardService.newCardSection().setHeader('Actions');
  
  actionSection.addWidget(
    CardService.newTextButton()
      .setText('Dismiss Job')
      .setOnClickAction(
        CardService.newAction()
          .setFunctionName('dismissJob')
          .setParameters({jobId: jobData.id})
      )
  );
  
  actionSection.addWidget(
    CardService.newTextButton()
      .setText('Back to Home')
      .setOnClickAction(CardService.newAction().setFunctionName('navigateToHome'))
  );
  
  builder.addSection(detailSection);
  builder.addSection(actionSection);
  
  return builder.build();
}

/**
 * Creates error display card
 */
function createErrorCard(errorMessage) {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Error'))
    .addSection(
      CardService.newCardSection()
        .addWidget(
          CardService.newTextParagraph().setText('An error occurred:')
        )
        .addWidget(
          CardService.newKeyValue()
            .setTopLabel("Details")
            .setContent(errorMessage)
            .setMultiline(true)
        )
    )
    .build();
}

// ================================================================================
// ACTION HANDLERS
// ================================================================================

/**
 * Navigation action to configuration card
 */
function navigateToConfigurationCard() {
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(createConfigurationCard()))
    .build();
}

/**
 * Navigation action back to home
 */
function navigateToHome() {
  const jobs = getActiveJobs();
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(createHomepageCard(jobs)))
    .build();
}

/**
 * Creates new report job from form input
 */
function createReportAction(e) {
  try {
    const formInputs = e.formInputs;
    
    // Validate required fields
    if (!formInputs.fullName || !formInputs.startDate || !formInputs.endDate) {
      return createNotification('Please fill in all required fields');
    }

    const selectedRole = formInputs.role_select ? formInputs.role_select[0] : 'INDIVIDUAL_CONTRIBUTOR';
    const weightsArray = ROLE_PRESETS[selectedRole].weights.split(',').map(w => parseInt(w.trim()));

    const userInputs = {
      fullName: formInputs.fullName[0],
      userEmail: formInputs.userEmail ? formInputs.userEmail[0] : Session.getActiveUser().getEmail(),
      notificationEmail: formInputs.notificationEmail ? formInputs.notificationEmail[0] : Session.getActiveUser().getEmail(),
      startDate: new Date(formInputs.startDate[0].msSinceEpoch),
      endDate: new Date(formInputs.endDate[0].msSinceEpoch),
      roleType: selectedRole,
      weights: {
        INNOVATION: weightsArray[0],
        EXECUTION: weightsArray[1],
        COLLABORATION: weightsArray[2], 
        LEADERSHIP: weightsArray[3]
      }
    };

    // Validate date range
    if (userInputs.startDate >= userInputs.endDate) {
      return createNotification('End date must be after start date');
    }

    // Create job and schedule processing
    const jobId = createJob(userInputs);
    scheduleJobProcessing(jobId);
    
    // Navigate back to home to show the new job
    const jobs = getActiveJobs();
    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().updateCard(createHomepageCard(jobs)))
      .setNotification(CardService.newNotification().setText('Report generation started'))
      .build();
      
  } catch (error) {
    console.error('Error in createReportAction:', error);
    return createNotification(`Error: ${error.toString()}`);
  }
}

/**
 * View job details action
 */
function viewJobDetails(e) {
  const jobId = e.parameters.jobId;
  const properties = PropertiesService.getUserProperties();
  const jobDataString = properties.getProperty(jobId);
  
  if (!jobDataString) {
    return createNotification("Job not found");
  }
  
  const jobData = JSON.parse(jobDataString);
  
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(createJobDetailCard(jobData)))
    .build();
}

/**
 * Open report and mark as viewed
 */
function openReportAndMarkViewed(e) {
  const jobId = e.parameters.jobId;
  const reportUrl = e.parameters.reportUrl;
  
  updateJobStatus(jobId, 'viewed');
  
  return CardService.newActionResponseBuilder()
    .setOpenLink(CardService.newOpenLink().setUrl(reportUrl))
    .setNotification(CardService.newNotification().setText('Report opened'))
    .build();
}

/**
 * Dismiss job action
 */
function dismissJob(e) {
  const jobId = e.parameters.jobId;
  removeJob(jobId);
  
  const jobs = getActiveJobs();
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(createHomepageCard(jobs)))
    .setNotification(CardService.newNotification().setText('Job dismissed'))
    .build();
}

/**
 * Creates notification response
 */
function createNotification(message) {
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText(message))
    .build();
}

// ================================================================================
// BACKGROUND PROCESSING
// ================================================================================

/**
 * Schedules job processing using time-based trigger
 */
function scheduleJobProcessing(jobId) {
  ScriptApp.newTrigger('processJob')
    .timeBased()
    .after(3 * 1000) // 3 seconds delay
    .create();
    
  // Store the job ID for the trigger to pick up
  PropertiesService.getScriptProperties().setProperty('currentJobId', jobId);
}

/**
 * Main background processing function updated for the new system
 */
function processJob() {
    var jobId = PropertiesService.getScriptProperties().getProperty('currentJobId');
    if (!jobId) {
        console.error('No job ID found for processing');
        return;
    }
    
    try {
        var jobDataString = PropertiesService.getUserProperties().getProperty(jobId);
        if (!jobDataString) {
            console.error('Job data not found:', jobId);
            return;
        }
        
        var jobData = JSON.parse(jobDataString);
        var userInputs = jobData.userInputs;
        
        // Convert date strings back to Date objects
        userInputs.startDate = new Date(userInputs.startDate);
        userInputs.endDate = new Date(userInputs.endDate);
        
        updateJobStatus(jobId, 'processing', 'Starting data collection...');
        
        // Fetch data
        updateJobStatus(jobId, 'processing', 'Fetching emails and calendar events...');
        var activities = fetchActivities(userInputs.userEmail, userInputs.startDate, userInputs.endDate);
        
        if (activities.length === 0) {
            throw new Error("No activities found in the selected date range");
        }
        
        updateJobStatus(jobId, 'processing', 'Analyzing ' + activities.length + ' activities for achievements and outcomes...');
        
        // Use the new achievement-focused analysis
        var analysis = analyzeAchievements(activities, userInputs.weights, userInputs);
        
        updateJobStatus(jobId, 'processing', 'Generating performance summary report...');
        
        // Generate report with new format
        var reportUrl = generateReport(analysis, userInputs);
        
        updateJobStatus(jobId, 'completed', 'Performance summary completed successfully', { reportUrl: reportUrl });
        
        // Send notification with updated metrics
        var summaryMetrics = {
            totalProjects: analysis.projectSummaries.length,
            strategicAlignment: analysis.executiveSummary.strategicAlignment,
            keyAchievements: analysis.executiveSummary.keyAchievements.length
        };
        
        sendCompletionNotification(userInputs.notificationEmail, userInputs.fullName, reportUrl, summaryMetrics);
        
    } catch (error) {
        console.error('Job processing failed:', error);
        updateJobStatus(jobId, 'failed', 'Processing failed', { errorMessage: error.toString() });
        
        // Try to get user inputs for error notification
        try {
            var jobDataString = PropertiesService.getUserProperties().getProperty(jobId);
            if (jobDataString) {
                var jobData = JSON.parse(jobDataString);
                sendErrorNotification(jobData.userInputs.notificationEmail, jobData.userInputs.fullName, error.toString());
            }
        } catch (notifyError) {
            console.error('Failed to send error notification:', notifyError);
        }
    } finally {
        // Clean up
        PropertiesService.getScriptProperties().deleteProperty('currentJobId');
        cleanupTriggers();
    }
}

/**
 * Updated completion notification for the new format
 */
function sendCompletionNotification(userEmail, fullName, reportUrl, metrics) {
    try {
        var subject = 'Performance Summary Report Ready - ' + fullName;
        var body = '<html><body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">' +
            '<div style="max-width: 600px; margin: 0 auto; padding: 20px;">' +
            '<h2 style="color: #00558C;">Your Performance Summary is Ready!</h2>' +
            '<p>Hello ' + fullName + ',</p>' +
            '<p>Your comprehensive performance summary has been generated, focusing on your key achievements, project contributions, and strategic impact.</p>' +
            
            '<div style="text-align: center; margin: 30px 0;">' +
            '<a href="' + reportUrl + '" ' +
            'style="background-color: #00558C; color: white; padding: 15px 30px; text-decoration: none; border-radius: 5px; font-weight: bold; display: inline-block;">' +
            'View Your Performance Summary' +
            '</a>' +
            '</div>' +
            
            '<div style="background-color: #f8f9fa; padding: 20px; border-radius: 8px; margin: 20px 0;">' +
            '<h3 style="margin-top: 0; color: #00558C;">Summary Highlights:</h3>' +
            '<ul style="margin-bottom: 0;">' +
            '<li><strong>Key Projects Analyzed:</strong> ' + metrics.totalProjects + '</li>' +
            '<li><strong>Strategic Alignment Score:</strong> ' + metrics.strategicAlignment.toFixed(1) + '/4.0</li>' +
            '<li><strong>Major Achievements:</strong> ' + metrics.keyAchievements + '</li>' +
            '</ul>' +
            '</div>' +
            
            '<p><strong>Your report includes:</strong></p>' +
            '<ul>' +
            '<li>Executive summary of your key achievements and strategic impact</li>' +
            '<li>Performance ratings across Innovation, Execution, Collaboration, and Leadership</li>' +
            '<li>Detailed project contributions showing your role and business outcomes</li>' +
            '<li>Strategic recommendations for continued growth and impact</li>' +
            '</ul>' +
            
            '<p style="margin-top: 30px;">Best regards,<br>Performance Summary Generator</p>' +
            '<hr style="margin-top: 30px; border: none; border-top: 1px solid #eee;">' +
            '<p style="font-size: 12px; color: #666;">This report focuses on your strategic contributions and business outcomes.</p>' +
            '</div>' +
            '</body></html>';

        GmailApp.sendEmail(userEmail, subject, '', {
            htmlBody: body,
            name: 'Performance Summary Generator'
        });
        
        console.log('Completion notification sent to: ' + userEmail);
    } catch (error) {
        console.error('Error sending completion notification:', error);
        throw error;
    }
}

/**
 * Cleans up old triggers
 */
function cleanupTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'processJob') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

// ================================================================================
// DATA FETCHING (IMPROVED)
// ================================================================================

/**
 * Fetches both email and calendar activities
 */
function fetchActivities(userEmail, startDate, endDate) {
  let activities = [];
  
  try {
    // Fetch calendar events
    const calendarActivities = fetchCalendarEvents(startDate, endDate);
    activities = activities.concat(calendarActivities);
    console.log(`Fetched ${calendarActivities.length} calendar events`);
    
    // Fetch email activities
    const emailActivities = fetchEmailActivities(userEmail, startDate, endDate);
    activities = activities.concat(emailActivities);
    console.log(`Fetched ${emailActivities.length} email activities`);
    
  } catch (error) {
    console.error('Error fetching activities:', error);
    throw error;
  }
  
  console.log(`Total activities fetched: ${activities.length}`);
  return activities;
}

/**
 * Fetches calendar events
 */
function fetchCalendarEvents(startDate, endDate) {
  const activities = [];
  
  try {
    const calendar = CalendarApp.getDefaultCalendar();
    const events = calendar.getEvents(startDate, endDate);
    
    events.forEach(event => {
      // Skip all-day events and very short meetings (likely not work-related)
      if (!event.isAllDayEvent() && 
          (event.getEndTime().getTime() - event.getStartTime().getTime()) >= 5 * 60 * 1000) {
        
        activities.push({
          type: 'Meeting',
          date: event.getStartTime(),
          title: event.getTitle() || 'Untitled Meeting',
          description: event.getDescription() || '',
          attendees: event.getGuestList().length,
          duration: (event.getEndTime().getTime() - event.getStartTime().getTime()) / (1000 * 60) // minutes
        });
      }
    });
    
  } catch (error) {
    console.error('Error fetching calendar events:', error);
    throw new Error(`Calendar fetch failed: ${error.toString()}`);
  }
  
  return activities;
}

/**
 * Fetches email activities with advanced Gmail search
 */
function fetchEmailActivities(userEmail, startDate, endDate) {
  const activities = [];
  
  try {
    // Build Gmail search query for better filtering
    const startDateStr = formatDateForGmail(startDate);
    const endDateStr = formatDateForGmail(endDate);
    
    // Search for sent emails
    const sentQuery = `from:${userEmail} after:${startDateStr} before:${endDateStr}`;
    const sentThreads = GmailApp.search(sentQuery, 0, 200);
    
    sentThreads.forEach(thread => {
      const messages = thread.getMessages();
      messages.forEach(message => {
        if (message.getFrom().includes(userEmail)) {
          const subject = message.getSubject();
          const body = message.getPlainBody();
          
          // Skip automated emails
          if (!isAutomatedEmail(subject, body)) {
            activities.push({
              type: 'Email',
              subtype: 'Sent',
              date: message.getDate(),
              title: subject || 'No Subject',
              description: body.substring(0, 500), // Limit description length
              recipients: message.getTo().split(',').length,
              threadLength: messages.length
            });
          }
        }
      });
    });
    
    // Search for important received emails (where user likely responded)
    const receivedQuery = `to:${userEmail} after:${startDateStr} before:${endDateStr} is:important`;
    const receivedThreads = GmailApp.search(receivedQuery, 0, 100);
    
    receivedThreads.forEach(thread => {
      const messages = thread.getMessages();
      // Only include if the user replied (thread has multiple messages and user sent one)
      const userReplied = messages.some(msg => msg.getFrom().includes(userEmail));
      
      if (userReplied && messages.length > 1) {
        const firstMessage = messages[0];
        const subject = firstMessage.getSubject();
        const body = firstMessage.getPlainBody();
        
        if (!isAutomatedEmail(subject, body)) {
          activities.push({
            type: 'Email',
            subtype: 'Important Thread',
            date: firstMessage.getDate(),
            title: subject || 'No Subject',
            description: body.substring(0, 500),
            recipients: firstMessage.getTo().split(',').length,
            threadLength: messages.length
          });
        }
      }
    });
    
  } catch (error) {
    console.error('Error fetching email activities:', error);
    throw new Error(`Email fetch failed: ${error.toString()}`);
  }
  
  return activities;
}

/**
 * Checks if an email is automated/system-generated
 */
function isAutomatedEmail(subject, body) {
  const automatedKeywords = [
    'noreply', 'no-reply', 'donotreply', 'automated', 'notification',
    'newsletter', 'unsubscribe', 'calendar', 'reminder', 'alert',
    'system generated', 'auto-generated'
  ];
  
  const lowercaseSubject = subject.toLowerCase();
  const lowercaseBody = body.toLowerCase();
  
  return automatedKeywords.some(keyword => 
    lowercaseSubject.includes(keyword) || lowercaseBody.includes(keyword)
  );
}

/**
 * Formats date for Gmail search
 */
function formatDateForGmail(date) {
  const year = date.getFullYear();
  const month = (date.getMonth() + 1).toString().padStart(2, '0');
  const day = date.getDate().toString().padStart(2, '0');
  return `${year}/${month}/${day}`;
}

// ================================================================================
// DATA ANALYSIS
// ================================================================================

/**
 * Analyzes activities and generates metrics
 */
function analyzeData(activities, weights) {
  const analysis = {
    distribution: {
      INNOVATION: { count: 0, totalScore: 0, items: [], rating: 0 },
      EXECUTION: { count: 0, totalScore: 0, items: [], rating: 0 },
      COLLABORATION: { count: 0, totalScore: 0, items: [], rating: 0 },
      LEADERSHIP: { count: 0, totalScore: 0, items: [], rating: 0 }
    },
    metrics: {
      totalActivities: 0,
      totalMeetings: 0,
      totalEmails: 0,
      averageRelevance: 0,
      weightedStrategicFocus: 0,
      highImpactActivities: 0
    },
    workPatterns: {
      byHour: Array(24).fill(0),
      byDay: Array(7).fill(0),
      communicationSplit: { internal: 0, external: 0 }
    },
    recommendations: []
  };

  if (!activities || activities.length === 0) {
    analysis.recommendations.push("No activities found to analyze. Try extending the date range or checking email permissions.");
    return analysis;
  }

  let totalRelevanceScore = 0;
  
  activities.forEach(activity => {
    analysis.metrics.totalActivities++;
    
    if (activity.type === 'Meeting') {
      analysis.metrics.totalMeetings++;
    } else if (activity.type === 'Email') {
      analysis.metrics.totalEmails++;
    }

    // Analyze text content for keywords
    const combinedText = `${activity.title} ${activity.description}`.toLowerCase();
    const scores = {};
    let bestCategory = null;
    let maxScore = 0;

    // Calculate keyword scores for each category
    for (const category in CONFIG.KEYWORDS) {
      scores[category] = CONFIG.SCORING.BASE_SCORE;
      
      CONFIG.KEYWORDS[category].forEach(keyword => {
        if (combinedText.includes(keyword.toLowerCase())) {
          scores[category] += CONFIG.SCORING.KEYWORD_HIT_VALUE;
        }
      });
      
      // Bonus scoring for collaboration indicators
      if (category === 'COLLABORATION') {
        if (activity.type === 'Meeting' && activity.attendees > 1) {
          scores[category] += CONFIG.SCORING.COLLABORATION_BONUS_MEETING;
        }
        if (activity.type === 'Email' && activity.recipients > 1) {
          scores[category] += CONFIG.SCORING.COLLABORATION_BONUS_EMAIL;
        }
      }
      
      // Track highest scoring category
      if (scores[category] > maxScore) {
        maxScore = scores[category];
        bestCategory = category;
      }
    }

    // Work pattern analysis
    const hour = new Date(activity.date).getHours();
    const day = new Date(activity.date).getDay();
    analysis.workPatterns.byHour[hour]++;
    analysis.workPatterns.byDay[day]++;

    // Assign activity to best matching category
    if (bestCategory && maxScore > CONFIG.SCORING.BASE_SCORE) {
      const relevanceScore = Math.min(maxScore, CONFIG.SCORING.MAX_SCORE);
      totalRelevanceScore += relevanceScore;
      activity.relevanceScore = relevanceScore;
      activity.category = bestCategory;

      analysis.distribution[bestCategory].count++;
      analysis.distribution[bestCategory].totalScore += relevanceScore;
      analysis.distribution[bestCategory].items.push(activity);

      // Track high impact activities
      if (relevanceScore >= 70) {
        analysis.metrics.highImpactActivities++;
      }
    }
  });

  // Calculate final metrics
  if (analysis.metrics.totalActivities > 0) {
    analysis.metrics.averageRelevance = totalRelevanceScore / analysis.metrics.totalActivities;
    
    // Calculate weighted strategic focus
    let weightedSum = 0;
    for (const category in analysis.distribution) {
      const percentageOfWork = (analysis.distribution[category].count / analysis.metrics.totalActivities) * 100;
      const weightValue = weights[category] || 0;
      weightedSum += (percentageOfWork * weightValue) / 100;
      
      // Calculate category ratings
      const avgScore = analysis.distribution[category].count > 0 
        ? analysis.distribution[category].totalScore / analysis.distribution[category].count 
        : 0;
        
      if (avgScore >= 80) analysis.distribution[category].rating = 4.0;
      else if (avgScore >= 65) analysis.distribution[category].rating = 3.0;
      else if (avgScore >= 45) analysis.distribution[category].rating = 2.0;
      else if (avgScore > 0) analysis.distribution[category].rating = 1.0;
      else analysis.distribution[category].rating = 0;
    }
    
    analysis.metrics.weightedStrategicFocus = weightedSum;
  }

  // Generate recommendations
  generateRecommendations(analysis);
  
  return analysis;
}

/**
 * Generates actionable recommendations based on analysis
 */
function generateRecommendations(analysis) {
  const totalActivities = analysis.metrics.totalActivities;
  
  if (totalActivities === 0) {
    analysis.recommendations.push("No activities found to analyze. Consider expanding the date range or checking permissions.");
    return;
  }

  // Calculate percentages for each category
  const percentages = {};
  const ratings = {};
  
  for (const category in analysis.distribution) {
    percentages[category] = (analysis.distribution[category].count / totalActivities) * 100;
    ratings[category] = analysis.distribution[category].rating;
  }

  // Innovation recommendations
  if (ratings.INNOVATION >= 3.5) {
    analysis.recommendations.push("🚀 INNOVATION EXCELLENCE: Consistently driving breakthrough solutions and competitive advantages.");
  } else if (percentages.INNOVATION > 25 && ratings.INNOVATION >= 2.5) {
    analysis.recommendations.push("💡 INNOVATION STRENGTH: Good focus on innovation. Document specific outcomes to maximize recognition.");
  } else if (percentages.INNOVATION < 15 || ratings.INNOVATION < 2.0) {
    analysis.recommendations.push("🔬 INNOVATION OPPORTUNITY: Increase time on R&D, prototyping, and creative problem-solving initiatives.");
  }

  // Execution recommendations
  if (ratings.EXECUTION >= 3.5) {
    analysis.recommendations.push("✅ EXECUTION MASTERY: Outstanding delivery record driving measurable business results.");
  } else if (percentages.EXECUTION > 30 && ratings.EXECUTION >= 2.5) {
    analysis.recommendations.push("🎯 EXECUTION STRENGTH: Strong delivery focus. Link activities to revenue/cost metrics for greater impact visibility.");
  } else if (percentages.EXECUTION < 20 || ratings.EXECUTION < 2.0) {
    analysis.recommendations.push("📈 EXECUTION FOCUS: Prioritize completion of key deliverables and milestone achievements.");
  }

  // Collaboration recommendations
  if (ratings.COLLABORATION >= 3.5) {
    analysis.recommendations.push("🤝 COLLABORATION LEADERSHIP: Exceptional cross-functional partnership driving organizational alignment.");
  } else if (percentages.COLLABORATION > 25 && ratings.COLLABORATION >= 2.5) {
    analysis.recommendations.push("🌐 COLLABORATION STRENGTH: Effective teamwork. Capture stakeholder feedback to quantify relationship impact.");
  } else if (percentages.COLLABORATION < 20 || ratings.COLLABORATION < 2.0) {
    analysis.recommendations.push("👥 COLLABORATION GROWTH: Increase cross-functional engagement and stakeholder partnership activities.");
  }

  // Leadership recommendations
  if (ratings.LEADERSHIP >= 3.5) {
    analysis.recommendations.push("👑 LEADERSHIP EXCELLENCE: Exceptional people development and strategic influence creating organizational capability.");
  } else if (percentages.LEADERSHIP > 15 && ratings.LEADERSHIP >= 2.5) {
    analysis.recommendations.push("🌟 LEADERSHIP IMPACT: Good leadership foundation. Track team performance improvements and strategic influence outcomes.");
  } else if (percentages.LEADERSHIP < 10 || ratings.LEADERSHIP < 2.0) {
    analysis.recommendations.push("🌱 LEADERSHIP DEVELOPMENT: Increase focus on mentoring, providing feedback, and strategic contributions to demonstrate leadership readiness.");
  }

  // Work pattern recommendations
  const outOfHoursActivities = analysis.workPatterns.byHour.slice(0, CONFIG.ANALYSIS.WORK_HOURS.START)
    .concat(analysis.workPatterns.byHour.slice(CONFIG.ANALYSIS.WORK_HOURS.END))
    .reduce((sum, count) => sum + count, 0);

  if (outOfHoursActivities > totalActivities * 0.25) {
    analysis.recommendations.push("⚖️ WORK-LIFE OPTIMIZATION: High after-hours activity detected. Consider delegation and boundary-setting for sustainable performance.");
  }

  // High-impact work recommendation
  const highImpactPercentage = (analysis.metrics.highImpactActivities / totalActivities) * 100;
  if (highImpactPercentage > 30) {
    analysis.recommendations.push("⭐ HIGH-IMPACT FOCUS: Excellent concentration on high-value activities. Continue prioritizing strategic work.");
  } else if (highImpactPercentage < 15) {
    analysis.recommendations.push("🎯 IMPACT OPTIMIZATION: Focus more time on high-value strategic activities. Consider delegating routine tasks.");
  }

  // Ensure we have at least one recommendation
  if (analysis.recommendations.length === 0) {
    analysis.recommendations.push("📊 BALANCED CONTRIBUTION: Well-rounded performance across key areas with opportunities for impact documentation.");
  }
}

// ================================================================================
// REPORT GENERATION
// ================================================================================

/**
 * Generates the Google Doc report
 */
function generateReport(analysis, userInputs) {
  let doc;
  
  try {
    const docName = `Work Summary - ${userInputs.fullName} (${userInputs.startDate.toLocaleDateString()})`;
    doc = DocumentApp.create(docName);
    const body = doc.getBody();
    body.clear();

    // Header section
    try {
      body.appendParagraph(doc.getName()).setAttributes(CONFIG.STYLE.TITLE);
      body.appendParagraph(`Analysis Period: ${userInputs.startDate.toLocaleDateString()} to ${userInputs.endDate.toLocaleDateString()}`)
        .setAttributes(CONFIG.STYLE.HEADER);
      body.appendParagraph(`Role Profile: ${ROLE_PRESETS[userInputs.roleType].name}`)
        .setAttributes(CONFIG.STYLE.HEADER);
      body.appendParagraph('');
    } catch (e) {
      throw new Error(`Failed at Header section: ${e.toString()}`);
    }

    // Executive Recommendations
    try {
      if (analysis.recommendations && analysis.recommendations.length > 0) {
        body.appendParagraph('Executive Recommendations').setAttributes(CONFIG.STYLE.H1);
        analysis.recommendations.forEach(rec => {
          body.appendListItem(rec);
        });
        body.appendParagraph('');
      }
    } catch (e) {
      throw new Error(`Failed at Recommendations section: ${e.toString()}`);
    }

    // Executive Summary
    try {
      body.appendParagraph('Executive Summary').setAttributes(CONFIG.STYLE.H1);
      insertExecutiveSummary(body, analysis.metrics);
    } catch (e) {
      throw new Error(`Failed at Executive Summary section: ${e.toString()}`);
    }

    // Key Values Analysis
    try {
      body.appendParagraph('Key Values Analysis').setAttributes(CONFIG.STYLE.H1);
      insertKeyValuesAnalysis(body, analysis.distribution, analysis.metrics.totalActivities, userInputs.weights);
    } catch (e) {
      throw new Error(`Failed at Key Values section: ${e.toString()}`);
    }

    // Work Patterns
    try {
      body.appendParagraph('Work Patterns').setAttributes(CONFIG.STYLE.H1);
      insertWorkPatterns(body, analysis.workPatterns, analysis.metrics);
    } catch (e) {
      throw new Error(`Failed at Work Patterns section: ${e.toString()}`);
    }

    // Top Activities
    try {
      body.appendParagraph('Top Strategic Activities').setAttributes(CONFIG.STYLE.H1);
      insertTopActivities(body, analysis);
    } catch (e) {
      throw new Error(`Failed at Top Activities section: ${e.toString()}`);
    }

    doc.saveAndClose();
    return doc.getUrl();
    
  } catch (error) {
    if (doc) {
      try {
        doc.getBody().appendParagraph(`\n\n--- REPORT GENERATION ERROR ---\n${error.toString()}`);
        doc.saveAndClose();
        return doc.getUrl();
      } catch (saveError) {
        console.error('Failed to save error document:', saveError);
      }
    }
    console.error('Error generating report:', error);
    throw new Error(`Failed to generate report: ${error.message}`);
  }
}

/**
 * Inserts executive summary table
 */
function insertExecutiveSummary(body, metrics) {
  const data = [
    ['Strategic Focus Score', `${metrics.weightedStrategicFocus.toFixed(1)}%`, 'Alignment with role priorities'],
    ['Average Impact Score', `${metrics.averageRelevance.toFixed(1)}/100`, 'Quality of work activities'],
    ['Total Activities', `${metrics.totalActivities}`, 'Meetings and emails analyzed'],
    ['High-Impact Activities', `${metrics.highImpactActivities}`, 'Activities with 70+ impact score']
  ];

  const table = body.appendTable(data);
  table.setBorderWidth(1);

  // Style the table
  for (let i = 0; i < data.length; i++) {
    table.getCell(i, 0).setAttributes(CONFIG.STYLE.TABLE_HEADER_CELL);
    table.getCell(i, 1).getChild(0).asParagraph().setAttributes(CONFIG.STYLE.GREEN_TEXT);
  }
  
  body.appendParagraph('');
}

/**
 * Inserts key values analysis
 */
function insertKeyValuesAnalysis(body, distribution, totalActivities, weights) {
  const categories = ['INNOVATION', 'EXECUTION', 'COLLABORATION', 'LEADERSHIP'];
  
  categories.forEach(category => {
    const data = distribution[category];
    const percentage = totalActivities > 0 ? ((data.count / totalActivities) * 100).toFixed(1) : '0.0';
    const avgScore = data.count > 0 ? (data.totalScore / data.count).toFixed(1) : '0.0';
    const weight = weights[category];
    
    body.appendParagraph(category).setAttributes(CONFIG.STYLE.H2);
    
    const summaryText = `Activities: ${data.count} (${percentage}% of total) | Average Score: ${avgScore}/100 | Rating: ${data.rating.toFixed(1)}/4.0 | Weight: ${weight}%`;
    body.appendParagraph(summaryText);
    
    body.appendParagraph('');
  });
}

/**
 * Inserts work patterns analysis
 */
function insertWorkPatterns(body, workPatterns, metrics) {
  body.appendParagraph('Communication Breakdown').setAttributes(CONFIG.STYLE.H2);
  body.appendParagraph(`Meetings: ${metrics.totalMeetings} | Emails: ${metrics.totalEmails}`);
  
  body.appendParagraph('Peak Activity Hours').setAttributes(CONFIG.STYLE.H2);
  const peakHour = workPatterns.byHour.indexOf(Math.max(...workPatterns.byHour));
  body.appendParagraph(`Most active hour: ${peakHour}:00 - ${peakHour + 1}:00`);
  
  body.appendParagraph('');
}

/**
 * Inserts top strategic activities
 */
function insertTopActivities(body, analysis) {
  const allActivities = [];
  
  // Collect all activities with scores
  for (const category in analysis.distribution) {
    analysis.distribution[category].items.forEach(item => {
      if (item.relevanceScore) {
        allActivities.push(item);
      }
    });
  }
  
  // Sort by relevance score and take top 10
  allActivities.sort((a, b) => b.relevanceScore - a.relevanceScore);
  const topActivities = allActivities.slice(0, 10);
  
  if (topActivities.length > 0) {
    topActivities.forEach((activity, index) => {
      body.appendParagraph(`${index + 1}. ${activity.title}`).setAttributes(CONFIG.STYLE.H2);
      body.appendParagraph(`Category: ${activity.category} | Score: ${activity.relevanceScore}/100 | Type: ${activity.type}`);
      if (activity.description && activity.description.length > 0) {
        body.appendParagraph(`Description: ${activity.description.substring(0, 200)}${activity.description.length > 200 ? '...' : ''}`);
      }
      body.appendParagraph('');
    });
  } else {
    body.appendParagraph('No high-scoring activities found. Consider reviewing the analysis period or expanding activity scope.');
  }
}

// ================================================================================
// NOTIFICATION SYSTEM
// ================================================================================

/**
 * Sends completion notification email
 */
function sendCompletionNotification(userEmail, fullName, reportUrl, metrics) {
  try {
    const subject = `✅ Your Work Summary Report is Ready - ${fullName}`;
    const body = `
<html>
<body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
  <div style="max-width: 600px; margin: 0 auto; padding: 20px;">
    <h2 style="color: #00558C;">🎉 Your Work Summary Report is Ready!</h2>
    <p>Hello ${fullName},</p>
    <p>Your comprehensive work summary report has been generated successfully and is ready for review.</p>
    
    <div style="text-align: center; margin: 30px 0;">
      <a href="${reportUrl}" 
         style="background-color: #00558C; color: white; padding: 15px 30px; text-decoration: none; border-radius: 5px; font-weight: bold; display: inline-block;">
        📄 Open Your Report
      </a>
    </div>
    
    <div style="background-color: #f8f9fa; padding: 20px; border-radius: 8px; margin: 20px 0;">
      <h3 style="margin-top: 0; color: #00558C;">Analysis Summary:</h3>
      <ul style="margin-bottom: 0;">
        <li><strong>Total Activities Analyzed:</strong> ${metrics.totalActivities}</li>
        <li><strong>Average Impact Score:</strong> ${metrics.averageRelevance.toFixed(1)}/100</li>
        <li><strong>High-Impact Activities:</strong> ${metrics.highImpactActivities}</li>
        <li><strong>Strategic Focus Score:</strong> ${metrics.weightedStrategicFocus.toFixed(1)}%</li>
      </ul>
    </div>
    
    <p><strong>What's included in your report:</strong></p>
    <ul>
      <li>Executive recommendations for career development</li>
      <li>Detailed analysis of your work across Innovation, Execution, Collaboration, and Leadership</li>
      <li>Work pattern insights and productivity metrics</li>
      <li>Top strategic activities with impact scores</li>
    </ul>
    
    <p style="margin-top: 30px;">Best regards,<br>Work Summary Generator</p>
    <hr style="margin-top: 30px; border: none; border-top: 1px solid #eee;">
    <p style="font-size: 12px; color: #666;">This is an automated notification. The report link will remain active and accessible through your Google Drive.</p>
  </div>
</body>
</html>
    `;

    GmailApp.sendEmail(userEmail, subject, '', {
      htmlBody: body,
      name: 'Work Summary Generator'
    });
    
    console.log(`Completion notification sent to: ${userEmail}`);
  } catch (error) {
    console.error('Error sending completion notification:', error);
    throw error;
  }
}

/**
 * Sends error notification email
 */
function sendErrorNotification(userEmail, fullName, errorMessage) {
  try {
    const subject = `❌ Work Summary Report Generation Failed - ${fullName}`;
    const body = `
<html>
<body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
  <div style="max-width: 600px; margin: 0 auto; padding: 20px;">
    <h2 style="color: #dc3545;">⚠️ Report Generation Failed</h2>
    <p>Hello ${fullName},</p>
    <p>Unfortunately, we encountered an issue while generating your work summary report.</p>
    
    <div style="background-color: #f8d7da; padding: 15px; border-radius: 5px; border-left: 4px solid #dc3545; margin: 20px 0;">
      <strong>Error Details:</strong><br>
      <code style="background-color: #fff; padding: 2px 4px; border-radius: 3px; font-family: monospace;">${errorMessage}</code>
    </div>
    
    <h3 style="color: #dc3545;">What you can do:</h3>
    <ol>
      <li><strong>Reduce the date range:</strong> Large date ranges can cause timeouts. Try analyzing 1-2 weeks at a time.</li>
      <li><strong>Check permissions:</strong> Ensure the add-on has access to Gmail and Google Drive.</li>
      <li><strong>Try again later:</strong> The issue might be temporary due to Google's API limits.</li>
      <li><strong>Contact support:</strong> If the problem persists, please reach out for assistance.</li>
    </ol>
    
    <div style="background-color: #d1ecf1; padding: 15px; border-radius: 5px; border-left: 4px solid #bee5eb; margin: 20px 0;">
      <strong>💡 Pro Tips:</strong>
      <ul style="margin-bottom: 0;">
        <li>Start with a 1-week analysis to test the system</li>
        <li>Ensure you have sent emails and calendar events in the selected period</li>
        <li>Check that your Gmail account has sufficient activity</li>
      </ul>
    </div>
    
    <p>We apologize for the inconvenience. Please don't hesitate to try again or contact us if you need assistance.</p>
    
    <p style="margin-top: 30px;">Best regards,<br>Work Summary Generator Support</p>
    <hr style="margin-top: 30px; border: none; border-top: 1px solid #eee;">
    <p style="font-size: 12px; color: #666;">This is an automated error notification.</p>
  </div>
</body>
</html>
    `;
    
    GmailApp.sendEmail(userEmail, subject, '', {
      htmlBody: body,
      name: 'Work Summary Generator'
    });
    
    console.log(`Error notification sent to: ${userEmail}`);
  } catch (error) {
    console.error('Error sending error notification:', error);
  }
}

// ================================================================================
// UTILITY AND TESTING FUNCTIONS
// ================================================================================

/**
 * Test function to validate permissions
 */
function testPermissions() {
  try {
    // Test Gmail access
    const aliases = GmailApp.getAliases();
    console.log('Gmail access: OK');
    
    // Test Calendar access
    const calendar = CalendarApp.getDefaultCalendar();
    console.log('Calendar access: OK');
    
    // Test Drive access
    const testDoc = DocumentApp.create('Permission Test - DELETE ME');
    const url = testDoc.getUrl();
    DriveApp.getFileById(testDoc.getId()).setTrashed(true);
    console.log('Drive access: OK');
    
    // Test Properties access
    PropertiesService.getUserProperties().setProperty('test', 'OK');
    const testValue = PropertiesService.getUserProperties().getProperty('test');
    PropertiesService.getUserProperties().deleteProperty('test');
    console.log('Properties access: OK');
    
    return 'All permissions validated successfully!';
  } catch (error) {
    console.error('Permission test failed:', error);
    throw new Error(`Permission test failed: ${error.toString()}`);
  }
}

/**
 * Cleanup function to remove old jobs and triggers
 */
function cleanupOldJobs(daysOld = 7) {
  const properties = PropertiesService.getUserProperties();
  const jobIdsString = properties.getProperty('activeJobIds');
  
  if (!jobIdsString) return;
  
  const jobIds = JSON.parse(jobIdsString);
  const cutoffDate = new Date(Date.now() - (daysOld * 24 * 60 * 60 * 1000));
  const activeJobs = [];
  
  jobIds.forEach(jobId => {
    const jobDataString = properties.getProperty(jobId);
    if (jobDataString) {
      const jobData = JSON.parse(jobDataString);
      const jobDate = new Date(jobData.startTime);
      
      if (jobDate > cutoffDate) {
        activeJobs.push(jobId);
      } else {
        // Remove old job
        properties.deleteProperty(jobId);
        console.log(`Cleaned up old job: ${jobId}`);
      }
    }
  });
  
  properties.setProperty('activeJobIds', JSON.stringify(activeJobs));
  console.log(`Cleanup complete. ${activeJobs.length} jobs remaining.`);
}

/**
 * Manual trigger to process a specific job (for debugging)
 */
function processSpecificJob(jobId) {
  PropertiesService.getScriptProperties().setProperty('currentJobId', jobId);
  processJob();
}
