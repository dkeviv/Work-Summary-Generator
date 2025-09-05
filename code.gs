// ================================================================================
// WORK SUMMARY GENERATOR - CLEAN SYNTAX VERSION
// ================================================================================

var CONFIG = {
    KEYWORDS: {
        INNOVATION: {
            stems: ['innovat', 'creat', 'design', 'prototype', 'pilot', 'experiment'],
            exact: ['mvp', 'r&d', 'poc', 'beta', 'alpha'],
            patterns: ['proof of concept', 'minimum viable', 'design thinking']
        },
        EXECUTION: {
            stems: ['deliver', 'ship', 'launch', 'complet', 'implement', 'deploy'],
            exact: ['go-live', 'rollout', 'milestone', 'deadline', 'kpi', 'roi'],
            patterns: ['project completion', 'goal completion', 'went live']
        },
        COLLABORATION: {
            stems: ['collaborat', 'partner', 'coordinat', 'align', 'sync', 'team'],
            exact: ['cross-functional', 'stakeholder', 'workshop', 'meeting'],
            patterns: ['working with', 'team effort', 'group effort']
        },
        LEADERSHIP: {
            stems: ['lead', 'mentor', 'coach', 'guid', 'direct', 'manag'],
            exact: ['1:1', 'one-on-one', 'feedback', 'delegation', 'vision'],
            patterns: ['decision making', 'strategic direction', 'team development']
        }
    },
    
    OUTCOME_KEYWORDS: {
        revenue: {
            stems: ['revenue', 'sales', 'profit', 'income', 'deal'],
            exact: ['roi', 'margin', 'pricing'],
            multiplier: 2.5
        },
        customer_satisfaction: {
            stems: ['customer', 'client', 'user', 'satisfaction', 'experience'],
            exact: ['nps', 'csat', 'churn', 'ux'],
            multiplier: 2.0
        },
        efficiency: {
            stems: ['efficiency', 'productivity', 'automat', 'optim'],
            exact: ['sla', 'turnaround'],
            multiplier: 1.8
        },
        market_position: {
            stems: ['market', 'competit', 'brand', 'growth'],
            exact: ['market share'],
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
        MAX_CATEGORY_SCORE: 100
    },

    STYLE: {
        HEADER: { [DocumentApp.Attribute.FONT_FAMILY]: 'Calibri', [DocumentApp.Attribute.FONT_SIZE]: 10 },
        TITLE: { [DocumentApp.Attribute.FONT_FAMILY]: 'Calibri', [DocumentApp.Attribute.FONT_SIZE]: 24, [DocumentApp.Attribute.BOLD]: true },
        H1: { [DocumentApp.Attribute.FONT_FAMILY]: 'Calibri', [DocumentApp.Attribute.FONT_SIZE]: 16, [DocumentApp.Attribute.BOLD]: true },
        H2: { [DocumentApp.Attribute.FONT_FAMILY]: 'Calibri', [DocumentApp.Attribute.FONT_SIZE]: 12, [DocumentApp.Attribute.BOLD]: true }
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
// JOB MANAGEMENT
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
    progress: 'Job created'
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
    .setHeader("Create New Report");
  
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
        .setContent(statusInfo.icon + ' ' + statusInfo.text);
        
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

function createConfigurationCard() {
  var currentUserEmail = Session.getActiveUser().getEmail();

  var builder = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Report Configuration'));

  var personalSection = CardService.newCardSection().setHeader('Personal Information');
  
  personalSection.addWidget(
    CardService.newTextInput()
      .setFieldName('fullName')
      .setTitle("Full Name")
  );

  personalSection.addWidget(
    CardService.newTextInput()
      .setFieldName('userEmail')
      .setTitle('Email to Analyze')
      .setValue(currentUserEmail)
  );

  personalSection.addWidget(
    CardService.newTextInput()
      .setFieldName('notificationEmail')
      .setTitle('Notification Email')
      .setValue(currentUserEmail)
  );

  builder.addSection(personalSection);

  var dateSection = CardService.newCardSection().setHeader('Analysis Period');
  
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

  var roleSection = CardService.newCardSection().setHeader('Role Profile');

  var roleSelect = CardService.newSelectionInput()
    .setFieldName('role_select')
    .setTitle("Role Type")
    .setType(CardService.SelectionInputType.DROPDOWN)
    .addItem(ROLE_PRESETS.INDIVIDUAL_CONTRIBUTOR.name, "INDIVIDUAL_CONTRIBUTOR", true)
    .addItem(ROLE_PRESETS.MANAGER.name, "MANAGER", false)
    .addItem(ROLE_PRESETS.EXECUTIVE.name, "EXECUTIVE", false);
  
  roleSection.addWidget(roleSelect);
  
  builder.addSection(roleSection);

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

function createErrorCard(errorMessage) {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Error'))
    .addSection(
      CardService.newCardSection()
        .addWidget(CardService.newTextParagraph().setText('An error occurred: ' + errorMessage))
    )
    .build();
}

// ================================================================================
// ACTION HANDLERS
// ================================================================================

function navigateToConfigurationCard() {
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(createConfigurationCard()))
    .build();
}

function createReportAction(e) {
  try {
    var formInputs = e.formInputs;
    
    if (!formInputs.fullName || !formInputs.startDate || !formInputs.endDate) {
      return createNotification('Please fill in all required fields');
    }

    var selectedRole = formInputs.role_select ? formInputs.role_select[0] : 'INDIVIDUAL_CONTRIBUTOR';
    var weightsArray = ROLE_PRESETS[selectedRole].weights.split(',');
    
    var userInputs = {
      fullName: formInputs.fullName[0],
      userEmail: formInputs.userEmail ? formInputs.userEmail[0] : Session.getActiveUser().getEmail(),
      notificationEmail: formInputs.notificationEmail ? formInputs.notificationEmail[0] : Session.getActiveUser().getEmail(),
      startDate: new Date(formInputs.startDate[0].msSinceEpoch),
      endDate: new Date(formInputs.endDate[0].msSinceEpoch),
      roleType: selectedRole,
      weights: {
        INNOVATION: parseInt(weightsArray[0]),
        EXECUTION: parseInt(weightsArray[1]),
        COLLABORATION: parseInt(weightsArray[2]),
        LEADERSHIP: parseInt(weightsArray[3])
      }
    };

    if (userInputs.startDate >= userInputs.endDate) {
      return createNotification('End date must be after start date');
    }

    var jobId = createJob(userInputs);
    scheduleJobProcessing(jobId);
    
    var jobs = getActiveJobs();
    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().updateCard(createHomepageCard(jobs)))
      .setNotification(CardService.newNotification().setText('Report generation started'))
      .build();
      
  } catch (error) {
    console.error('Error in createReportAction:', error);
    return createNotification('Error: ' + error.toString());
  }
}

function viewJobDetails(e) {
  var jobId = e.parameters.jobId;
  var properties = PropertiesService.getUserProperties();
  var jobDataString = properties.getProperty(jobId);
  
  if (!jobDataString) {
    return createNotification("Job not found");
  }
  
  var jobData = JSON.parse(jobDataString);
  
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(createJobDetailCard(jobData)))
    .build();
}

function createJobDetailCard(jobData) {
  var builder = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Job Details'));

  var detailSection = CardService.newCardSection();
  
  var statusInfo = getStatusInfo(jobData.status);
  
  detailSection.addWidget(
    CardService.newKeyValue()
      .setTopLabel('Name')
      .setContent(jobData.userInputs.fullName)
  );
  
  detailSection.addWidget(
    CardService.newKeyValue()
      .setTopLabel('Status')
      .setContent(statusInfo.icon + ' ' + statusInfo.text)
  );
  
  if (jobData.reportUrl) {
    detailSection.addWidget(
      CardService.newTextButton()
        .setText('Open Report')
        .setOnClickAction(
          CardService.newAction()
            .setFunctionName('openReportAndMarkViewed')
            .setParameters({jobId: jobData.id, reportUrl: jobData.reportUrl})
        )
    );
  }
  
  if (jobData.errorMessage) {
    detailSection.addWidget(
      CardService.newKeyValue()
        .setTopLabel('Error')
        .setContent(jobData.errorMessage)
    );
  }
  
  detailSection.addWidget(
    CardService.newTextButton()
      .setText('Dismiss Job')
      .setOnClickAction(
        CardService.newAction()
          .setFunctionName('dismissJob')
          .setParameters({jobId: jobData.id})
      )
  );
  
  builder.addSection(detailSection);
  
  return builder.build();
}

function openReportAndMarkViewed(e) {
  var jobId = e.parameters.jobId;
  var reportUrl = e.parameters.reportUrl;
  
  updateJobStatus(jobId, 'viewed');
  
  return CardService.newActionResponseBuilder()
    .setOpenLink(CardService.newOpenLink().setUrl(reportUrl))
    .build();
}

function dismissJob(e) {
  var jobId = e.parameters.jobId;
  removeJob(jobId);
  
  var jobs = getActiveJobs();
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(createHomepageCard(jobs)))
    .build();
}

function createNotification(message) {
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText(message))
    .build();
}

// ================================================================================
// BACKGROUND PROCESSING
// ================================================================================

function scheduleJobProcessing(jobId) {
  ScriptApp.newTrigger('processJob')
    .timeBased()
    .after(3 * 1000)
    .create();
    
  PropertiesService.getScriptProperties().setProperty('currentJobId', jobId);
}

function processJob() {
  var jobId = PropertiesService.getScriptProperties().getProperty('currentJobId');
  if (!jobId) {
    console.error('No job ID found');
    return;
  }
  
  try {
    var jobDataString = PropertiesService.getUserProperties().getProperty(jobId);
    if (!jobDataString) {
      console.error('Job data not found');
      return;
    }
    
    var jobData = JSON.parse(jobDataString);
    var userInputs = jobData.userInputs;
    
    userInputs.startDate = new Date(userInputs.startDate);
    userInputs.endDate = new Date(userInputs.endDate);
    
    updateJobStatus(jobId, 'processing', 'Fetching data...');
    
    var activities = fetchActivities(userInputs.userEmail, userInputs.startDate, userInputs.endDate);
    
    if (activities.length === 0) {
      throw new Error("No activities found");
    }
    
    updateJobStatus(jobId, 'processing', 'Analyzing activities...');
    
    var analysis = analyzeAchievements(activities, userInputs.weights, userInputs);
    
    updateJobStatus(jobId, 'processing', 'Generating report...');
    
    var reportUrl = generateReport(analysis, userInputs);
    
    updateJobStatus(jobId, 'completed', 'Report completed', { reportUrl: reportUrl });
    
    sendCompletionNotification(userInputs.notificationEmail, userInputs.fullName, reportUrl);
    
  } catch (error) {
    console.error('Job failed:', error);
    updateJobStatus(jobId, 'failed', 'Failed', { errorMessage: error.toString() });
  } finally {
    PropertiesService.getScriptProperties().deleteProperty('currentJobId');
    cleanupTriggers();
  }
}

function cleanupTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var trigger = triggers[i];
    if (trigger.getHandlerFunction() === 'processJob') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
}

// ================================================================================
// SIMPLIFIED ANALYSIS
// ================================================================================

function analyzeAchievements(activities, weights, userInputs) {
  var analysis = {
    executiveSummary: {
      totalImpact: 0,
      keyAchievements: [],
      strategicAlignment: 0
    },
    performanceMetrics: {
      INNOVATION: { rating: 0, impact: 0, description: '' },
      EXECUTION: { rating: 0, impact: 0, description: '' },
      COLLABORATION: { rating: 0, impact: 0, description: '' },
      LEADERSHIP: { rating: 0, impact: 0, description: '' }
    },
    projectSummaries: [],
    recommendations: []
  };

  if (!activities || activities.length === 0) {
    analysis.recommendations.push("No activities found. Try extending the date range.");
    return analysis;
  }

  // Simplified project grouping
  var projects = groupActivitiesIntoProjects(activities);
  analysis.projectSummaries = analyzeProjectContributions(projects, userInputs);
  analysis.performanceMetrics = calculatePerformanceMetrics(analysis.projectSummaries);
  analysis.executiveSummary = generateExecutiveSummary(analysis.projectSummaries);
  analysis.recommendations = generateRecommendations(analysis.performanceMetrics);
  
  return analysis;
}

function calculateActivityScore(text) {
  var score = CONFIG.SCORING.BASE_SCORE;
  var labels = [];
  var leadershipMultiplier = 1.0;
  var outcomeMultiplier = 1.0;
  var contributionType = 'Contributed';
  var businessOutcomes = [];

  // Improved fuzzy/stem/exact/pattern matching and multi-labeling
  function matchKeywords(keywordObj, category) {
    var matched = false;
    keywordObj.stems.forEach(function(stem) {
      if (text.match(new RegExp(stem, 'i'))) {
        score += CONFIG.SCORING.STEM_MATCH_VALUE;
        matched = true;
        if (labels.indexOf(category) === -1) labels.push(category);
      }
    });
    keywordObj.exact.forEach(function(exact) {
      if (text.match(new RegExp('\\b' + exact + '\\b', 'i'))) {
        score += CONFIG.SCORING.EXACT_MATCH_VALUE;
        matched = true;
        if (labels.indexOf(category) === -1) labels.push(category);
      }
    });
    keywordObj.patterns && keywordObj.patterns.forEach(function(pattern) {
      if (text.toLowerCase().indexOf(pattern.toLowerCase()) >= 0) {
        score += CONFIG.SCORING.PATTERN_MATCH_VALUE;
        matched = true;
        if (labels.indexOf(category) === -1) labels.push(category);
      }
    });
    return matched;
  }

  Object.keys(CONFIG.KEYWORDS).forEach(function(category) {
    matchKeywords(CONFIG.KEYWORDS[category], category);
  });

  // Leadership detection and weighting
  if (matchKeywords(CONFIG.KEYWORDS.LEADERSHIP)) {
    leadershipMultiplier = 2.5;
    if (text.match(/stepped in|filled in|backup|strategic|company-wide|long-term vision|mentor|coaching|team coverage/i)) {
      score += 20;
      labels.push('Leadership Enhancement');
    }
  }

  // Outcome multipliers
  Object.keys(CONFIG.OUTCOME_KEYWORDS).forEach(function(outcome) {
    var ok = CONFIG.OUTCOME_KEYWORDS[outcome];
    if (matchKeywords(ok)) {
      outcomeMultiplier = Math.max(outcomeMultiplier, ok.multiplier);
      businessOutcomes.push(outcome);
    }
  });

  // Contribution analysis
  if (text.match(/led|lead|drove|managed|initiated/i)) {
    contributionType = 'Led';
    score += 15;
  } else if (text.match(/collaborat|partnered|coordinated|worked with/i)) {
    contributionType = 'Collaborated';
    score += CONFIG.SCORING.COLLABORATION_BONUS;
  } else if (text.match(/contributed|supported|assisted/i)) {
    contributionType = 'Contributed';
    score += 5;
  }

  // Apply multipliers
  score = score * leadershipMultiplier * outcomeMultiplier;
  score = Math.min(score, CONFIG.SCORING.MAX_CATEGORY_SCORE);

  return {
    score: score,
    labels: labels,
    leadershipMultiplier: leadershipMultiplier,
    outcomeMultiplier: outcomeMultiplier,
    contributionType: contributionType,
    businessOutcomes: businessOutcomes
  };
}

function groupActivitiesIntoProjects(activities) {
  // Improved clustering: group by most common phrase/topic in activities
  var clusters = {};
  var titleMap = {};
  for (var i = 0; i < activities.length; i++) {
    var activity = activities[i];
    var combinedText = (activity.title + ' ' + activity.description).toLowerCase();
    var scoreObj = calculateActivityScore(combinedText);
    activity.impactScore = scoreObj.score;
    activity.labels = scoreObj.labels;
    activity.contributionType = scoreObj.contributionType;
    activity.businessOutcomes = scoreObj.businessOutcomes;
    // Use most frequent label as cluster key
    var key = scoreObj.labels.length > 0 ? scoreObj.labels.join('-') : 'Other';
    if (!clusters[key]) clusters[key] = [];
    clusters[key].push(activity);
    // Track most common title
    var t = activity.title || 'Untitled';
    titleMap[key] = titleMap[key] || {};
    titleMap[key][t] = (titleMap[key][t]||0)+1;
  }
  // Create projects from clusters
  var projects = [];
  Object.keys(clusters).forEach(function(key) {
    var cluster = clusters[key];
    var totalImpact = cluster.reduce(function(sum, act) { return sum + act.impactScore; }, 0);
    // Normalize totalImpact to 1-100 scale
    var normImpact = Math.max(1, Math.min(Math.round(totalImpact / cluster.length), 100));
    var timeframe = {
      start: cluster[0].date,
      end: cluster[cluster.length-1].date
    };
    // Use most frequent title for project name
    var titleCounts = titleMap[key];
    var topTitle = Object.keys(titleCounts).sort(function(a,b){return titleCounts[b]-titleCounts[a];})[0] || 'Untitled';
    // Use all labels for multi-labeling
    var allLabels = [];
    cluster.forEach(function(act){act.labels.forEach(function(lab){if(allLabels.indexOf(lab)===-1)allLabels.push(lab);});});
    projects.push({
      title: cleanTitle(topTitle),
      activities: cluster,
      totalImpact: normImpact,
      labels: allLabels,
      timeframe: timeframe
    });
  });
  projects.sort(function(a, b) { return b.totalImpact - a.totalImpact; });
  return projects.slice(0, 10);
}

function analyzeProjectContributions(projects, userInputs) {
  var summaries = [];
  for (var i = 0; i < projects.length; i++) {
    var project = projects[i];
    var mainActivity = project.activities[0];
    var contributionType = mainActivity.contributionType;
    var businessOutcomes = mainActivity.businessOutcomes;
    var outcomeDescriptions = [];
    var outcomeNarrative = '';
    if (businessOutcomes.indexOf('customer_satisfaction') >= 0) {
      outcomeDescriptions.push('Customer Impact: Improved experience or satisfaction');
      outcomeNarrative += 'This project delivered measurable improvements in customer experience and satisfaction, aligning with our commitment to customer-centricity. ';
    }
    if (businessOutcomes.indexOf('revenue') >= 0) {
      outcomeDescriptions.push('Business Impact: Revenue growth or margin improvement');
      outcomeNarrative += 'The initiative contributed directly to revenue growth and margin improvement, supporting organizational financial goals. ';
    }
    if (businessOutcomes.indexOf('efficiency') >= 0) {
      outcomeDescriptions.push('Efficiency: Process optimization or automation');
      outcomeNarrative += 'Efforts in this project led to process optimization and automation, driving operational efficiency and cost savings. ';
    }
    if (businessOutcomes.indexOf('market_position') >= 0) {
      outcomeDescriptions.push('Market Impact: Enhanced market position or share');
      outcomeNarrative += 'Strategic actions taken improved our market position and competitive standing. ';
    }
    // Executive narrative - more substantive, connects to outcomes and org goals
    var narrative = userInputs.fullName + ' ' + contributionType + ' the project "' + project.title + '" (' + project.activities.length + ' activities | Impact Score: ' + project.totalImpact.toFixed(0) + ').\n';
    narrative += 'Role: ' + contributionType + ' | Collaboration: ' + (project.activities.length > 5 ? 'Large group collaboration (8+ people)' : 'Small team') + '.\n';
    narrative += 'Key achievements include: ';
    if (outcomeDescriptions.length > 0) {
      narrative += outcomeDescriptions.join('; ') + '. ';
    }
    narrative += outcomeNarrative;
    narrative += 'This project advanced organizational goals such as revenue growth, customer satisfaction, efficiency, and market position. '; 
    narrative += 'Specific activities included: ';
    for (var a = 0; a < project.activities.length; a++) {
      var act = project.activities[a];
      narrative += '\n- ' + act.title + ': ' + (act.description ? act.description.substring(0,80) : '') + '...';
    }
    narrative += '\nOverall, this work is a strong example of achievement-focused execution, directly supporting key business outcomes.';
    var summary = {
      projectTitle: project.title,
      totalImpact: project.totalImpact,
      contributionType: contributionType,
      businessOutcomes: outcomeDescriptions,
      narrative: narrative
    };
    summaries.push(summary);
  }
  return summaries;
}

function calculatePerformanceMetrics(projectSummaries) {
    // Increase ratings by 20 percentage points (multiply by 1.2, max 4.0)
      // Dynamic rating calculation based on projectSummaries
      var categories = ['INNOVATION', 'EXECUTION', 'COLLABORATION', 'LEADERSHIP'];
      var metrics = {};
      var maxImpact = 0;
      var categoryImpact = { INNOVATION: 0, EXECUTION: 0, COLLABORATION: 0, LEADERSHIP: 0 };
      var categoryCounts = { INNOVATION: 0, EXECUTION: 0, COLLABORATION: 0, LEADERSHIP: 0 };

      // Aggregate impact scores by category
      for (var i = 0; i < projectSummaries.length; i++) {
        var project = projectSummaries[i];
        var labels = project.labels || [];
        var impact = project.totalImpact || 0;
        maxImpact = Math.max(maxImpact, impact);
        for (var j = 0; j < labels.length; j++) {
          var label = labels[j];
          if (categoryImpact[label] !== undefined) {
            categoryImpact[label] += impact;
            categoryCounts[label] += 1;
          }
        }
      }

      // If no impact, set maxImpact to 1 to avoid division by zero
      if (maxImpact === 0) maxImpact = 1;

      // Calculate ratings for each category
      categories.forEach(function(category) {
        var avgImpact = categoryCounts[category] > 0 ? categoryImpact[category] / categoryCounts[category] : 0;
        // Normalize to 1.0–4.0 scale, ensure spread
        var rating = 1.0 + (avgImpact / 100) * 3.0;
        rating = Math.min(Math.max(rating, 1.0), 4.0);
        metrics[category] = {
          rating: rating,
          impact: categoryImpact[category],
          description: categoryCounts[category] > 0 ? 'Significant contributions in ' + category.toLowerCase() : 'Limited activity in ' + category.toLowerCase()
        };
      });
      return metrics;
}

function generateExecutiveSummary(projectSummaries) {
  var summary = {
    totalImpact: 0,
    keyAchievements: [],
    strategicAlignment: 2.7
  };
  
  for (var i = 0; i < Math.min(3, projectSummaries.length); i++) {
    var project = projectSummaries[i];
    summary.totalImpact += project.totalImpact;
    
    summary.keyAchievements.push({
      title: project.projectTitle,
      contribution: project.contributionType,
      impact: project.totalImpact,
      outcomes: ['Business Impact']
    });
  }
  
  return summary;
}

function generateRecommendations(performanceMetrics) {
  var recommendations = [];
  // Portfolio analysis based recommendations
  if (performanceMetrics.EXECUTION.rating > 3.0) {
    recommendations.push('STRENGTH - EXECUTION: Maintain delivery excellence and scale best practices.');
  } else {
    recommendations.push('OPPORTUNITY - EXECUTION: Focus on improving delivery consistency.');
  }
  if (performanceMetrics.INNOVATION.rating < 3.0) {
    recommendations.push('OPPORTUNITY - INNOVATION: Increase creative problem-solving and pilot new ideas.');
  } else {
    recommendations.push('STRENGTH - INNOVATION: Continue driving innovation and experimentation.');
  }
  if (performanceMetrics.LEADERSHIP.rating < 3.0) {
    recommendations.push('DEVELOPING - LEADERSHIP: Take on more mentoring, team coverage, and strategic responsibilities.');
  } else {
    recommendations.push('STRENGTH - LEADERSHIP: Demonstrated strong leadership and organizational thinking.');
  }
  if (performanceMetrics.COLLABORATION.rating < 3.0) {
    recommendations.push('OPPORTUNITY - COLLABORATION: Expand cross-functional collaboration and stakeholder engagement.');
  } else {
    recommendations.push('STRENGTH - COLLABORATION: Effective cross-functional teamwork.');
  }
  return recommendations;
}

// ================================================================================
// DATA FETCHING
// ================================================================================

function fetchActivities(userEmail, startDate, endDate) {
  var activities = [];
  
  try {
    var calendarActivities = fetchCalendarEvents(startDate, endDate);
    activities = activities.concat(calendarActivities);
    
    var emailActivities = fetchEmailActivities(userEmail, startDate, endDate);
    activities = activities.concat(emailActivities);
    
  } catch (error) {
    console.error('Error fetching activities:', error);
    throw error;
  }
  
  return activities;
}

function fetchCalendarEvents(startDate, endDate) {
  var activities = [];
  
  try {
    var calendar = CalendarApp.getDefaultCalendar();
    var events = calendar.getEvents(startDate, endDate);
    
    for (var i = 0; i < events.length; i++) {
      var event = events[i];
      if (!event.isAllDayEvent()) {
        activities.push({
          type: 'Meeting',
          date: event.getStartTime(),
          title: event.getTitle() || 'Meeting',
          description: event.getDescription() || '',
          attendees: event.getGuestList().length
        });
      }
    }
    
  } catch (error) {
    console.error('Error fetching calendar:', error);
  }
  
  return activities;
}

function fetchEmailActivities(userEmail, startDate, endDate) {
  var activities = [];
  
  try {
    var startDateStr = formatDateForGmail(startDate);
    var endDateStr = formatDateForGmail(endDate);
    
    var query = 'from:' + userEmail + ' after:' + startDateStr + ' before:' + endDateStr;
    var threads = GmailApp.search(query, 0, 100);
    
    for (var i = 0; i < threads.length; i++) {
      var thread = threads[i];
      var messages = thread.getMessages();
      for (var j = 0; j < messages.length; j++) {
        var message = messages[j];
        if (message.getFrom().indexOf(userEmail) >= 0) {
          var subject = message.getSubject();
          var body = message.getPlainBody();
          
          if (!isAutomatedEmail(subject, body)) {
            activities.push({
              type: 'Email',
              date: message.getDate(),
              title: subject || 'No Subject',
              description: body.substring(0, 300),
              recipients: message.getTo().split(',').length
            });
          }
        }
      }
    }
    
  } catch (error) {
    console.error('Error fetching emails:', error);
  }
  
  return activities;
}

function isAutomatedEmail(subject, body) {
  var automatedKeywords = ['noreply', 'no-reply', 'automated', 'notification', 'newsletter'];
  var lowercaseSubject = subject.toLowerCase();
  var lowercaseBody = body.toLowerCase();
  
  for (var i = 0; i < automatedKeywords.length; i++) {
    var keyword = automatedKeywords[i];
    if (lowercaseSubject.indexOf(keyword) >= 0 || lowercaseBody.indexOf(keyword) >= 0) {
      return true;
    }
  }
  
  return false;
}

function formatDateForGmail(date) {
  var year = date.getFullYear();
  var month = (date.getMonth() + 1).toString();
  if (month.length < 2) month = '0' + month;
  var day = date.getDate().toString();
  if (day.length < 2) day = '0' + day;
  return year + '/' + month + '/' + day;
}

// ================================================================================
// REPORT GENERATION
// ================================================================================

function generateReport(analysis, userInputs) {
  var doc;
  
  try {
    var docName = 'Performance Summary - ' + userInputs.fullName + ' (' + userInputs.startDate.toLocaleDateString() + ')';
    doc = DocumentApp.create(docName);
    var body = doc.getBody();
    body.clear();

    body.appendParagraph(docName).setAttributes(CONFIG.STYLE.TITLE);
    body.appendParagraph('Analysis Period: ' + userInputs.startDate.toLocaleDateString() + ' to ' + userInputs.endDate.toLocaleDateString())
        .setAttributes(CONFIG.STYLE.HEADER);
    body.appendParagraph('');

    body.appendParagraph('Executive Summary').setAttributes(CONFIG.STYLE.H1);
    insertExecutiveSummary(body, analysis.executiveSummary);

    body.appendParagraph('Performance Metrics').setAttributes(CONFIG.STYLE.H1);
    insertPerformanceMetrics(body, analysis.performanceMetrics);

    body.appendParagraph('Key Project Contributions').setAttributes(CONFIG.STYLE.H1);
    insertProjectContributions(body, analysis.projectSummaries);

    body.appendParagraph('Strategic Recommendations').setAttributes(CONFIG.STYLE.H1);
    for (var i = 0; i < analysis.recommendations.length; i++) {
      body.appendListItem(analysis.recommendations[i]);
    }

    doc.saveAndClose();
    return doc.getUrl();
    
  } catch (error) {
    if (doc) {
      try {
        doc.getBody().appendParagraph('\\n\\nREPORT ERROR: ' + error.toString());
        doc.saveAndClose();
        return doc.getUrl();
      } catch (saveError) {
        console.error('Failed to save error document:', saveError);
      }
    }
    console.error('Error generating report:', error);
    throw new Error('Failed to generate report: ' + error.message);
  }
}

function insertExecutiveSummary(body, executiveSummary) {
  if (executiveSummary.keyAchievements.length > 0) {
    body.appendParagraph('Key Achievements:').setAttributes(CONFIG.STYLE.H2);
    for (var i = 0; i < executiveSummary.keyAchievements.length; i++) {
      var achievement = executiveSummary.keyAchievements[i];
      var achievementText = achievement.contribution + ' ' + achievement.title;
      if (achievement.outcomes.length > 0) {
        achievementText += ' (Impact: ' + achievement.outcomes.join(', ') + ')';
      }
      body.appendListItem(achievementText);
    }
    body.appendParagraph('');
  }
  
  body.appendParagraph('Strategic Alignment: ' + executiveSummary.strategicAlignment.toFixed(1) + '/4.0')
      .setAttributes(CONFIG.STYLE.H2);
  body.appendParagraph('');
}

function insertPerformanceMetrics(body, performanceMetrics) {
  var categories = ['INNOVATION', 'EXECUTION', 'COLLABORATION', 'LEADERSHIP'];
  
  for (var i = 0; i < categories.length; i++) {
    var category = categories[i];
    var metric = performanceMetrics[category];
    
    if (metric.rating > 0) {
      body.appendParagraph(category).setAttributes(CONFIG.STYLE.H2);
      body.appendParagraph('Rating: ' + metric.rating.toFixed(1) + '/4.0 | Impact Score: ' + metric.impact.toFixed(0));
      if (metric.description) {
        body.appendParagraph('Key Contributions: ' + metric.description);
      }
      body.appendParagraph('');
    }
  }
}

function insertProjectContributions(body, projectSummaries) {
  for (var i = 0; i < projectSummaries.length; i++) {
    var project = projectSummaries[i];
    body.appendParagraph(project.projectTitle).setAttributes(CONFIG.STYLE.H2);
    body.appendParagraph('Impact Score: ' + project.totalImpact.toFixed(0));
    body.appendParagraph(project.narrative);
    if (project.businessOutcomes.length > 0) {
      body.appendParagraph('Business Outcomes:').setAttributes(CONFIG.STYLE.H2);
      for (var j = 0; j < project.businessOutcomes.length; j++) {
        body.appendListItem(project.businessOutcomes[j]);
      }
    }
    body.appendParagraph('');
  }
}

// ================================================================================
// NOTIFICATION SYSTEM
// ================================================================================

function sendCompletionNotification(userEmail, fullName, reportUrl) {
  try {
    var subject = 'Performance Summary Ready - ' + fullName;
    var body = 'Hello ' + fullName + ',\n\n' +
        'Your performance summary report is ready.\n\n' +
        'View your report: ' + reportUrl + '\n\n' +
        'Best regards,\n' +
        'Performance Summary Generator';

    GmailApp.sendEmail(userEmail, subject, body);
    console.log('Notification sent to: ' + userEmail);
  } catch (error) {
    console.error('Error sending notification:', error);
  }
}

function sendErrorNotification(userEmail, fullName, errorMessage) {
  try {
    var subject = 'Performance Summary Failed - ' + fullName;
    var body = 'Hello ' + fullName + ',\n\n' +
        'Your performance summary generation failed.\n\n' +
        'Error: ' + errorMessage + '\n\n' +
        'Please try again with a smaller date range.\n\n' +
        'Best regards,\n' +
        'Performance Summary Generator';
    
    GmailApp.sendEmail(userEmail, subject, body);
    console.log('Error notification sent to: ' + userEmail);
  } catch (error) {
    console.error('Error sending error notification:', error);
  }
}

// ================================================================================
// UTILITY FUNCTIONS
// ================================================================================

function testPermissions() {
  try {
    GmailApp.getAliases();
    console.log('Gmail access: OK');
    
    CalendarApp.getDefaultCalendar();
    console.log('Calendar access: OK');
    
    var testDoc = DocumentApp.create('Test Doc');
    DriveApp.getFileById(testDoc.getId()).setTrashed(true);
    console.log('Drive access: OK');
    
    PropertiesService.getUserProperties().setProperty('test', 'OK');
    PropertiesService.getUserProperties().deleteProperty('test');
    console.log('Properties access: OK');
    
    return 'All permissions validated successfully!';
  } catch (error) {
    console.error('Permission test failed:', error);
    throw new Error('Permission test failed: ' + error.toString());
  }
}

function cleanupOldJobs(daysOld) {
  if (!daysOld) daysOld = 7;
  
  var properties = PropertiesService.getUserProperties();
  var jobIdsString = properties.getProperty('activeJobIds');
  
  if (!jobIdsString) return;
  
  var jobIds = JSON.parse(jobIdsString);
  var cutoffDate = new Date(Date.now() - (daysOld * 24 * 60 * 60 * 1000));
  var activeJobs = [];
  
  for (var i = 0; i < jobIds.length; i++) {
    var jobId = jobIds[i];
    var jobDataString = properties.getProperty(jobId);
    if (jobDataString) {
      var jobData = JSON.parse(jobDataString);
      var jobDate = new Date(jobData.startTime);
      
      if (jobDate > cutoffDate) {
        activeJobs.push(jobId);
      } else {
        properties.deleteProperty(jobId);
        console.log('Cleaned up old job: ' + jobId);
      }
    }
  }
  
  properties.setProperty('activeJobIds', JSON.stringify(activeJobs));
  console.log('Cleanup complete. ' + activeJobs.length + ' jobs remaining.');
}

function cleanTitle(title) {
  if (!title) return 'Untitled Project';
  title = title.replace(/^(re:|fwd:|fw:)/gi, '');
  title = title.replace(/^[\s-]*/, '');
  if (title.length > 0) {
    title = title.charAt(0).toUpperCase() + title.slice(1);
  }
  return title.substring(0, 60);
}

