// --- GLOBAL CONFIGURATION ---
const CONFIG = {
    KEYWORDS: { STRATEGIC_PLANNING: ['strategy', 'strategic', 'planning', 'roadmap', 'qbr', 'okr', 'goals', 'vision', 'budget', 'forecast', 'initiative'], LEADERSHIP_COMMS: ['1:1', 'one-on-one', 'team sync', 'all-hands', 'mentoring', 'coaching', 'feedback', 'review', 'hiring', 'headcount', 'development', 'update', 'internal comms'], OPERATIONAL_EXECUTION: ['support', 'ticket', 'bug', 'fix', 'deploy', 'release', 'ops', 'maintenance', 'troubleshoot', 'incident', 'report', 'standup', 'daily sync'], CLIENT_STAKEHOLDER: ['client:', 'customer:', 'partner:', 'q&a', 'demo', 'proposal', 'contract', 'negotiation', 'feedback from', 'follow-up with']},
    SCORING: { BASE_SCORE: 10, KEYWORD_HIT_VALUE: 15, COLLABORATION_BONUS_MEETING: 20, COLLABORATION_BONUS_EMAIL: 10, EXTERNAL_COMM_BONUS: 15, MAX_SCORE: 100},
    STYLE: { HEADER: {[DocumentApp.Attribute.FONT_FAMILY]: 'Calibri', [DocumentApp.Attribute.FONT_SIZE]: 10, [DocumentApp.Attribute.ITALIC]: true}, TITLE: {[DocumentApp.Attribute.FONT_FAMILY]: 'Calibri', [DocumentApp.Attribute.FONT_SIZE]: 24, [DocumentApp.Attribute.BOLD]: true}, H1: {[DocumentApp.Attribute.FONT_FAMILY]: 'Calibri', [DocumentApp.Attribute.FONT_SIZE]: 16, [DocumentApp.Attribute.BOLD]: true, [DocumentApp.Attribute.FOREGROUND_COLOR]: '#00558C'}, H2: {[DocumentApp.Attribute.FONT_FAMILY]: 'Calibri', [DocumentApp.Attribute.FONT_SIZE]: 12, [DocumentApp.Attribute.BOLD]: true}, TABLE_HEADER_CELL: {[DocumentApp.Attribute.BOLD]: true, [DocumentApp.Attribute.BACKGROUND_COLOR]: '#F3F3F3'}, GREEN_TEXT: {[DocumentApp.Attribute.FOREGROUND_COLOR]: '#38761D'}, ORANGE_TEXT: {[DocumentApp.Attribute.FOREGROUND_COLOR]: '#B45F06'}, RED_TEXT: {[DocumentApp.Attribute.FOREGROUND_COLOR]: '#990000'}},
    ANALYSIS: { TOP_N_ACTIVITIES: 5, MAX_GMAIL_THREADS: 250, WORK_HOURS: {START: 8, END: 18}}
};

const ROLE_PRESETS = {
  "EXECUTIVE": "50,35,5,10",
  "MANAGER": "20,40,25,15",
  "INDIVIDUAL_CONTRIBUTOR": "10,15,60,15"
};

/**
 * The trigger function that builds the initial UI when the add-on is opened in Gmail.
 */
function onGmailHomepage(e) {
  return createWelcomeCard();
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
    return CardService.newActionResponseBuilder()
        .setNavigation(CardService.newNavigation().pushCard(createConfigurationCard()))
        .build();
}

/**
 * Creates the main UI card with all the user-friendly input fields.
 * @param {Object} e The event object, used to repopulate fields.
 */
function createConfigurationCard(e = {}) {
  const formInputs = e.formInputs || {};
  const selectedRole = formInputs.role_select ? formInputs.role_select[0] : '';
  const weights = selectedRole ? ROLE_PRESETS[selectedRole] : (formInputs.weights ? formInputs.weights[0] : '');

  const card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader().setTitle('Create New Report').setSubtitle('Step 1: Configure Your Report'));

  // --- Section 1: Who and When ---
  const whoWhenSection = CardService.newCardSection().setHeader('Who & When');
  
  whoWhenSection.addWidget(CardService.newTextInput().setFieldName('fullName').setTitle("Person's Full Name"));
  whoWhenSection.addWidget(CardService.newTextInput().setFieldName('userEmail').setTitle('Email to Analyze').setValue(Session.getActiveUser().getEmail()));
  whoWhenSection.addWidget(CardService.newTextParagraph().setText("<i><b>Note:</b> For security, this script can only analyze the data of the person running it.</i>"));
  whoWhenSection.addWidget(CardService.newDatePicker().setFieldName('startDate').setTitle('Start Date'));
  whoWhenSection.addWidget(CardService.newDatePicker().setFieldName('endDate').setTitle('End Date'));
  
  card.addSection(whoWhenSection);

  // --- Section 2 for Topical Summary ---
  const topicSection = CardService.newCardSection().setHeader('Topical Deep Dive (Optional)');
  topicSection.addWidget(CardService.newTextInput()
      .setFieldName('topic')
      .setTitle('Summarize a topic')
      .setHint("e.g., 'Project Phoenix' or 'Q4 Budget'"));
      
  topicSection.addWidget(CardService.newTextInput()
      .setFieldName('personEmail')
      .setTitle('Filter topic by person (Optional)')
      .setHint("e.g., jane.doe@example.com"));

  card.addSection(topicSection);


  // --- Section 3: Strategy Weights ---
  const weightsSection = CardService.newCardSection().setHeader('Define Role Focus (Strategy Weights)');
  
  weightsSection.addWidget(CardService.newTextParagraph()
    .setText("This tells the tool what's most important for the role. Think of it like a <b>budget for professional focus.</b> Start with a preset and adjust if needed."));

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
  
  weightsSection.addWidget(CardService.newTextInput()
    .setFieldName('weights')
    .setTitle('Strategy Weights (Strategic, Leadership, Operational, Client)')
    .setValue(weights)
    .setHint('Must be 4 numbers separated by commas, totaling 100'));

  card.addSection(weightsSection);

  // --- Section 4: Generate Button ---
  card.setFixedFooter(CardService.newFixedFooter().setPrimaryButton(
      CardService.newTextButton()
        .setText('Generate Report')
        .setOnClickAction(CardService.newAction().setFunctionName('createReportAction'))
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)));

  return card.build();
}

/**
 * Action handler that re-renders the configuration card when a role preset is chosen.
 */
function handleRoleChange(e) {
    return CardService.newActionResponseBuilder()
        .setNavigation(CardService.newNavigation().updateCard(createConfigurationCard(e)))
        .build();
}

/**
 * The action handler function that is called when the "Generate Report" button is clicked.
 */
function createReportAction(e) {
  const formInputs = e.formInputs;
  const userEmail = formInputs.userEmail[0];
  const fullName = formInputs.fullName[0];
  const startDateMs = formInputs.startDate ? formInputs.startDate[0].msSinceEpoch : null;
  const endDateMs = formInputs.endDate ? formInputs.endDate[0].msSinceEpoch : null;
  const weightsStr = formInputs.weights[0];
  const topic = formInputs.topic ? formInputs.topic[0].trim() : null;
  const personEmail = formInputs.personEmail ? formInputs.personEmail[0].trim() : null;

  if (!fullName || !startDateMs || !endDateMs || !weightsStr) {
    return CardService.newActionResponseBuilder().setNotification(CardService.newNotification().setText('Error: All fields in "Who & When" and "Strategy Weights" are required.')).build();
  }
  
  const userInputs = {
      userEmail: userEmail, fullName: fullName, startDate: new Date(startDateMs), endDate: new Date(endDateMs),
      userDomain: userEmail.split('@')[1], weights: {}, topic: topic, personEmail: personEmail
  };

  const weightsArr = weightsStr.split(',').map(w => parseInt(w.trim(), 10));
  if (weightsArr.length !== 4 || weightsArr.some(isNaN) || weightsArr.reduce((a, b) => a + b, 0) !== 100) {
     return CardService.newActionResponseBuilder().setNotification(CardService.newNotification().setText('Error: Invalid weights. Must be 4 numbers that sum to 100.')).build();
  }
  userInputs.weights = { STRATEGIC_PLANNING: weightsArr[0], LEADERSHIP_COMMS: weightsArr[1], OPERATIONAL_EXECUTION: weightsArr[2], CLIENT_STAKEHOLDER: weightsArr[3] };
  
  try {
    const rawData = fetchData(userInputs.startDate, userInputs.endDate, userInputs.userEmail);
    const analysis = analyzeData(rawData, userInputs.weights, userInputs.userDomain, userInputs.topic, userInputs.personEmail);
    const docUrl = generateReport(analysis, userInputs);

    const successCard = createSuccessCard(docUrl, analysis.metrics);
    return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().pushCard(successCard)).build();

  } catch (err) {
     const errorCard = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle('An Error Occurred').setSubtitle('Please try again.'))
      .addSection(CardService.newCardSection().addWidget(CardService.newTextParagraph().setText(err.message))).build();
    return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().pushCard(errorCard)).build();
  }
}

/**
 * Creates the final success card with a link to the report.
 */
function createSuccessCard(docUrl, metrics) {
    return CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle('Report Generation Complete!'))
      .addSection(CardService.newCardSection()
        .addWidget(CardService.newKeyValue().setTopLabel('Status').setContent('✅ Success').setIcon(CardService.Icon.CONFIRMATION))
        .addWidget(CardService.newKeyValue().setTopLabel('Strategic Focus Score').setContent(`${metrics.weightedStrategicFocus.toFixed(1)}%`))
        .addWidget(CardService.newKeyValue().setTopLabel('Activities Analyzed').setContent(metrics.totalActivities.toString()))
      )
      .addSection(CardService.newCardSection()
          .addWidget(CardService.newButtonSet()
              .addButton(CardService.newTextButton().setText('Open Report').setOpenLink(CardService.newOpenLink().setUrl(docUrl)).setTextButtonStyle(CardService.TextButtonStyle.FILLED))
              .addButton(CardService.newTextButton().setText('Create Another').setOnClickAction(CardService.newAction().setFunctionName('handleCreateAnother')))
          )
      ).build();
}

/**
 * Action handler to go back to the welcome screen.
 */
function handleCreateAnother() {
    return CardService.newActionResponseBuilder()
        .setNavigation(CardService.newNavigation().popToRoot())
        .build();
}


// =================================================================================
// == CORE LOGIC FUNCTIONS                                                      ==
// =================================================================================

function fetchData(startDate, endDate, userEmail) {
  if (userEmail !== Session.getActiveUser().getEmail()) {
    throw new Error(`Security restriction: This script can only analyze the data for ${Session.getActiveUser().getEmail()}, the user running it.`);
  }
  let activities = [];
  const events = CalendarApp.getDefaultCalendar().getEvents(startDate, endDate);
  for (const event of events) {
    const userStatus = event.getGuestByEmail(userEmail)?.getGuestStatus();
    if (userStatus === CalendarApp.GuestStatus.YES || event.isOwnedBy(userEmail)) {
      activities.push({type: 'Meeting', id: event.getId(), date: event.getStartTime(), title: event.getTitle(), description: event.getDescription(), attendees: event.getGuestList().map(g => g.getEmail())});
    }
  }
  const query = `from:${userEmail} after:${startDate.toISOString().slice(0, 10)} before:${endDate.toISOString().slice(0, 10)}`;
  const threads = GmailApp.search(query, 0, CONFIG.ANALYSIS.MAX_GMAIL_THREADS);
  for (const thread of threads) {
    const message = thread.getMessages()[0];
    const recipients = [message.getTo(), message.getCc()].join(',').split(',').filter(e => e && e.trim() !== '');
    activities.push({type: 'Email', id: message.getId(), date: message.getDate(), title: message.getSubject(), description: message.getPlainBody().substring(0, 500), attendees: recipients});
  }
  return activities;
}

/**
 * Main analysis function, now includes topic summarization by a specific person.
 */
function analyzeData(activities, weights, userDomain, topic, personEmail) {
  const analysis = {
    activities: [], metrics: {}, distribution: { STRATEGIC_PLANNING: {count: 0, totalScore: 0, items: []}, LEADERSHIP_COMMS: {count: 0, totalScore: 0, items: []}, OPERATIONAL_EXECUTION: {count: 0, totalScore: 0, items: []}, CLIENT_STAKEHOLDER: {count: 0, totalScore: 0, items: []}, UNCATEGORIZED: {count: 0, totalScore: 0, items: []}}, workPatterns: {byHour: Array(24).fill(0), commSplit: {internal: 0, external: 0}}, recommendations: [], topicSummary: null };
  if (activities.length === 0) {
    analysis.metrics = { totalActivities: 0, totalMeetings: 0, totalEmails: 0, averageRelevance: 0, weightedStrategicFocus: 0 };
    analysis.recommendations.push("No activities found for the selected period.");
    return analysis; 
  }

  // Run topic summary if a topic is provided
  if (topic && topic.length > 0) {
    analysis.topicSummary = summarizeTopicFromActivities(topic, personEmail, activities);
  }

  // Perform the main work categorization and scoring
  for (const activity of activities) {
    const textCorpus = `${activity.title.toLowerCase()} ${activity.description.toLowerCase()}`;
    let bestCategory = 'UNCATEGORIZED';
    let maxHits = 0;
    for (const category in CONFIG.KEYWORDS) {
      const hits = CONFIG.KEYWORDS[category].filter(kw => textCorpus.includes(kw)).length;
      if (hits > maxHits) { maxHits = hits; bestCategory = category; }
    }
    activity.category = bestCategory;
    let score = CONFIG.SCORING.BASE_SCORE + (maxHits * CONFIG.SCORING.KEYWORD_HIT_VALUE);
    if (activity.type === 'Meeting' && activity.attendees.length > 1) score += CONFIG.SCORING.COLLABORATION_BONUS_MEETING;
    if (activity.type === 'Email' && activity.attendees.length > 1) score += CONFIG.SCORING.COLLABORATION_BONUS_EMAIL;
    const isExternal = activity.attendees.some(a => a && a.includes('@') && !a.endsWith('@' + userDomain));
    if (isExternal) score += CONFIG.SCORING.EXTERNAL_COMM_BONUS;
    activity.relevanceScore = Math.min(CONFIG.SCORING.MAX_SCORE, Math.round(score));
    analysis.activities.push(activity);
    const dist = analysis.distribution[bestCategory];
    dist.count++; dist.totalScore += activity.relevanceScore; dist.items.push(activity);
    analysis.workPatterns.byHour[activity.date.getHours()]++;
    isExternal ? analysis.workPatterns.commSplit.external++ : analysis.workPatterns.commSplit.internal++;
  }
  const totalActivities = activities.length;
  analysis.metrics.totalActivities = totalActivities;
  analysis.metrics.totalMeetings = activities.filter(a => a.type === 'Meeting').length;
  analysis.metrics.totalEmails = activities.filter(a => a.type === 'Email').length;
  analysis.metrics.averageRelevance = analysis.activities.reduce((sum, a) => sum + a.relevanceScore, 0) / totalActivities || 0;
  let weightedScoreSum = 0;
  const totalWeight = Object.values(weights).reduce((s,w) => s+w, 0) || 100;
  for (const category in analysis.distribution) {
    if (weights[category]) {
      const categoryPercentage = (analysis.distribution[category].count / totalActivities) * 100;
      weightedScoreSum += categoryPercentage * (weights[category] / totalWeight);
    }
  }
  analysis.metrics.weightedStrategicFocus = weightedScoreSum;
  const strategicPct = (analysis.distribution.STRATEGIC_PLANNING.count / totalActivities) * 100;
  if (strategicPct > 40) analysis.recommendations.push("✅ Strong focus on strategic initiatives is evident.");
  if ((analysis.distribution.OPERATIONAL_EXECUTION.count / totalActivities) * 100 > 50) analysis.recommendations.push("📈 Significant time on operational tasks. Consider delegation opportunities.");
  if (analysis.recommendations.length === 0) analysis.recommendations.push("📊 Work distribution appears balanced.");
  return analysis;
}

/**
 * Finds and summarizes emails related to a topic, with an optional filter for a specific person.
 */
function summarizeTopicFromActivities(topic, personEmail, activities) {
  const lowerCaseTopic = topic.toLowerCase();
  
  // First, filter by the topic
  let relevantEmails = activities.filter(activity => 
    activity.type === 'Email' &&
    (activity.title.toLowerCase().includes(lowerCaseTopic) || 
     activity.description.toLowerCase().includes(lowerCaseTopic))
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
    email.attendees.forEach(p => uniqueParticipants.add(p));
  });

  return {
    count: relevantEmails.length,
    topSubjects: relevantEmails.slice(0, 5).map(email => email.title),
    participants: Array.from(uniqueParticipants)
  };
}


function generateReport(analysis, inputs) {
  const docName = `Work Summary - ${inputs.fullName} (${inputs.startDate.toISOString().slice(0, 10)})`;
  const doc = DocumentApp.create(docName);
  const body = doc.getBody();
  body.appendParagraph(doc.getName()).setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph(`Analysis Period: ${inputs.startDate.toLocaleDateString()} to ${inputs.endDate.toLocaleDateString()}`);
  
  // Insert the topic summary section if it exists
  if (analysis.topicSummary) {
    insertTopicSummary(body, analysis.topicSummary, inputs.topic, inputs.personEmail);
  }
  
  insertExecutiveSummary(body, analysis.metrics);
  body.appendParagraph('Performance Metrics').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  insertMetricsTable(body, analysis.metrics);
  body.appendParagraph('Activity Distribution').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  insertDistributionTable(body, analysis.distribution, analysis.metrics.totalActivities);
  body.appendParagraph('Detailed Activities').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  insertDetailedActivities(body, analysis.distribution);
  body.appendParagraph('Work Patterns').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  insertWorkPatterns(body, analysis.workPatterns, analysis.metrics.totalActivities);
  body.appendParagraph('Executive Recommendations').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  analysis.recommendations.forEach(rec => body.appendListItem(rec));
  doc.saveAndClose();
  return doc.getUrl();
}

/**
 * Writes the topic summary to the Google Doc, now with a dynamic title.
 */
function insertTopicSummary(body, topicSummary, topic, personEmail) {
  let heading = 'Topical Deep Dive: "' + topic + '"';
  if (personEmail && personEmail.length > 0) {
    heading += ' with ' + personEmail;
  }
  
  body.appendParagraph(heading).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  const tableData = [
    ['Metric', 'Details'],
    ['Matching Email Threads', topicSummary.count.toString()],
    ['Key Conversation Subjects', topicSummary.topSubjects.join('\n') || 'N/A'],
    ['Key Participants', topicSummary.participants.slice(0, 10).join(', ') + (topicSummary.participants.length > 10 ? '...' : '')]
  ];
  createStyledTable(body, tableData);
}

function insertExecutiveSummary(body, metrics) {
  const data = [['Strategic Focus Score', `${metrics.weightedStrategicFocus.toFixed(1)}%`, 'Alignment with priorities.'], ['Avg. Relevance Score', `${metrics.averageRelevance.toFixed(1)} / 100`, 'Average impact of activities.'], ['Total Activities', `${metrics.totalActivities}`, `Meetings & sent emails.`]];
  const table = body.appendTable(data);
  table.setBorderWidth(0);
  for (let i = 0; i < data.length; i++) {
    table.getCell(i, 0).setAttributes(CONFIG.STYLE.H2);
    table.getCell(i, 1).getChild(0).asParagraph().setAttributes(CONFIG.STYLE.GREEN_TEXT).setBold(true).setFontSize(14);
    table.getCell(i, 2).getChild(0).asParagraph().setItalic(true);
  }
  body.appendParagraph('');
}

function insertMetricsTable(body, metrics) {
  const data = [['Metric', 'Value'], ['Total Activities', metrics.totalActivities], ['Meetings', metrics.totalMeetings], ['Emails Sent', metrics.totalEmails], ['Avg. Relevance', `${metrics.averageRelevance.toFixed(1)}%`]];
  createStyledTable(body, data);
}

function insertDistributionTable(body, distribution, total) {
  if (total === 0) { body.appendParagraph("No activities found to distribute."); return; };
  const data = [['Category', '% of Activities', 'Avg. Relevance']];
  const map = {STRATEGIC_PLANNING: '🗺️ Strategic', LEADERSHIP_COMMS: '👥 Leadership', OPERATIONAL_EXECUTION: '⚙️ Operational', CLIENT_STAKEHOLDER: '🤝 Client-Facing', UNCATEGORIZED: '📥 Uncategorized'};
  for (const cat in map) {
    const item = distribution[cat];
    if (item.count > 0) { data.push([map[cat], `${((item.count/total)*100).toFixed(1)}%`, `${(item.totalScore/item.count).toFixed(1)}%`]); }
  }
  createStyledTable(body, data);
}

function insertDetailedActivities(body, distribution) {
  const map = {STRATEGIC_PLANNING: '🗺️ Top Strategic Activities', LEADERSHIP_COMMS: '👥 Top Leadership Activities', OPERATIONAL_EXECUTION: '⚙️ Top Operational Activities', CLIENT_STAKEHOLDER: '🤝 Top Client-Facing Activities'};
  let hasActivities = false;
  for (const cat in map) {
    const items = distribution[cat].items;
    if (items.length > 0) {
      hasActivities = true;
      body.appendParagraph(map[cat]).setHeading(DocumentApp.ParagraphHeading.HEADING2);
      const sorted = items.sort((a,b)=>b.relevanceScore - a.relevanceScore).slice(0, CONFIG.ANALYSIS.TOP_N_ACTIVITIES);
      sorted.forEach(item => {
        const style = item.relevanceScore > 70 ? CONFIG.STYLE.GREEN_TEXT : item.relevanceScore > 40 ? CONFIG.STYLE.ORANGE_TEXT : CONFIG.STYLE.RED_TEXT;
        const li = body.appendListItem(`[${item.date.toLocaleDateString()}] ${item.title}`);
        li.appendText(` (Score: `).appendText(`${item.relevanceScore}`).setAttributes(style).setBold(true).appendText(`)`);
      });
    }
  }
  if (!hasActivities) body.appendParagraph("No categorized activities to detail.");
}

function insertWorkPatterns(body, patterns, total) {
  if (total === 0) { body.appendParagraph("No work patterns to analyze."); return; };
  const totalComm = patterns.commSplit.internal + patterns.commSplit.external;
  body.appendParagraph('Communication Focus').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendListItem(`Internal: ${totalComm > 0 ? ((patterns.commSplit.internal/totalComm)*100).toFixed(0) : 0}%`);
  body.appendListItem(`External: ${totalComm > 0 ? ((patterns.commSplit.external/totalComm)*100).toFixed(0) : 0}%`);
  body.appendParagraph('');
  let outOfHours = 0;
  patterns.byHour.forEach((count, hour) => { if (count > 0 && (hour < CONFIG.ANALYSIS.WORK_HOURS.START || hour >= CONFIG.ANALYSIS.WORK_HOURS.END)) { outOfHours += count; } });
  body.appendParagraph('Activity Timing').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendListItem(`Work outside standard hours: ${total > 0 ? ((outOfHours/total)*100).toFixed(0) : 0}%`);
  body.appendListItem(`Peak hour: ${patterns.byHour.indexOf(Math.max(...patterns.byHour))}:00`);
}

function createStyledTable(body, data) {
  if (data.length < 2) { body.appendParagraph("No data available."); return; }
  const table = body.appendTable(data);
  table.getRow(0).setAttributes(CONFIG.STYLE.TABLE_HEADER_CELL);
  body.appendParagraph('');
}
