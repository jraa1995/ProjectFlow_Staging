var BUILTIN_DOC_TEMPLATES = {
  designDoc: {
    label: 'Design Documentation',
    sections: [
      { heading: 'Overview & Purpose', body: 'Describe what this project is and why it exists. What problem does it solve?' },
      { heading: 'Goals & Success Criteria', body: 'List measurable goals and acceptance criteria.' },
      { heading: 'Users & Stakeholders', body: 'Who are the primary users? Who has decision authority?' },
      { heading: 'User Stories / Use Cases', body: 'Document the key user stories or use cases this project supports.' },
      { heading: 'High-Level Architecture', body: 'Sketch the high-level architecture or include a link to a diagram.' },
      { heading: 'UI / UX Design', body: 'Describe key screens, mockups, or design system references.' },
      { heading: 'Data Flow', body: 'Trace how data moves through the system.' },
      { heading: 'Technical Constraints', body: 'List browser/runtime/API constraints, security and compliance requirements.' },
      { heading: 'Open Questions / Risks', body: 'List unresolved questions, risks, and assumptions.' },
      { heading: 'Decision Log', body: 'Record significant design decisions with rationale and date.' }
    ]
  },
  devDoc: {
    label: 'Development Documentation',
    sections: [
      { heading: 'Project Overview', body: 'Brief project description, status, and current phase.' },
      { heading: 'Tech Stack', body: 'List primary languages, frameworks, libraries, and runtime versions.' },
      { heading: 'Repository / Source Code', body: 'Link to the Git repository and branching strategy.' },
      { heading: 'Architecture & Components', body: 'Describe the major components/modules and how they interact.' },
      { heading: 'Setup & Local Development', body: 'Steps a new developer takes to clone, install, and run locally.' },
      { heading: 'Environment Variables / Configuration', body: 'List required environment variables and their purpose.' },
      { heading: 'Build & Deployment', body: 'How to build and deploy. Include CI/CD pipeline references.' },
      { heading: 'Key Workflows', body: 'Document important code paths (e.g., authentication, data ingest, scheduled jobs).' },
      { heading: 'Testing', body: 'How to run tests. Coverage targets. Critical test scenarios.' },
      { heading: 'Known Issues / Tech Debt', body: 'Track outstanding issues, workarounds, and tech debt.' }
    ]
  },
  dataDictionary: {
    label: 'Data Dictionary',
    sections: [
      { heading: 'Overview', body: 'Summary of the data managed by this project and its scope.' },
      { heading: 'Data Sources', body: 'List source systems, ingestion methods, and ownership.' },
      { heading: 'Tables / Datasets / Files', body: 'For each dataset: name, location, format, refresh cadence, row count.' },
      { heading: 'Field Definitions', body: 'For each field: name, data type, allowed values, source, sensitivity.' },
      { heading: 'Refresh Cadence', body: 'How often each dataset is refreshed and via what process.' },
      { heading: 'Data Owners / Stewards', body: 'Primary contact for data questions and access requests.' },
      { heading: 'Lineage', body: 'Trace how raw data is transformed into the final output.' },
      { heading: 'Quality & Validation Rules', body: 'Rules used to validate data quality and detect anomalies.' },
      { heading: 'Retention & Compliance', body: 'Retention policy, PII handling, and compliance notes.' }
    ]
  },
  userGuide: {
    label: 'User Guide',
    sections: [
      { heading: 'Overview', body: 'What this project does for the user, in plain language.' },
      { heading: 'Getting Started', body: 'How to access the application or report. Required permissions.' },
      { heading: 'Common Tasks', body: 'Step-by-step instructions for the most-used workflows.' },
      { heading: 'Tips & Best Practices', body: 'Guidance to help users get the most value.' },
      { heading: 'Troubleshooting', body: 'Common issues users may encounter and how to resolve them.' },
      { heading: 'FAQs', body: 'Frequently asked questions and short answers.' },
      { heading: 'Support & Feedback', body: 'How to report issues, request features, or get help.' }
    ]
  },
  supportRunbook: {
    label: 'Support Runbook',
    sections: [
      { heading: 'Project Overview', body: 'High-level summary of what this project does and who it serves.' },
      { heading: 'Architecture Quick-Reference', body: 'Where the code lives, where it runs, what it depends on.' },
      { heading: 'On-Call Escalation Path', body: 'Primary owner, backup owner, and escalation contacts.' },
      { heading: 'Common Failure Modes', body: 'Known failure scenarios and how to detect them.' },
      { heading: 'Standard Resolutions', body: 'For each common failure, the documented fix.' },
      { heading: 'Monitoring & Alerts', body: 'Where alerts fire. Health-check links and thresholds.' },
      { heading: 'Recovery Procedures', body: 'Steps to recover from outages or data corruption.' },
      { heading: 'Useful Commands & Links', body: 'Quick-reference commands, dashboards, and tickets.' },
      { heading: 'Change Log', body: 'Significant operational changes with date and reason.' }
    ]
  }
};

function getBuiltinDocTemplates() {
  try {
    var summary = {};
    Object.keys(BUILTIN_DOC_TEMPLATES).forEach(function(key) {
      var t = BUILTIN_DOC_TEMPLATES[key];
      summary[key] = {
        key: key,
        label: t.label,
        sectionCount: t.sections.length
      };
    });
    return { success: true, templates: summary };
  } catch (error) {
    console.error('getBuiltinDocTemplates failed:', error);
    return { success: false, error: error.message };
  }
}

function generateDocFromBuiltinTemplate(projectId, templateKey) {
  try {
    if (!projectId) throw new Error('projectId is required');
    if (!templateKey) throw new Error('templateKey is required');
    var template = BUILTIN_DOC_TEMPLATES[templateKey];
    if (!template) throw new Error('Unknown template: ' + templateKey);

    PermissionGuard.requirePermission('project:update', { projectId: projectId });

    var project = getProjectById(projectId);
    if (!project) throw new Error('Project not found: ' + projectId);

    var settings = {};
    try { settings = typeof project.settings === 'string' ? JSON.parse(project.settings || '{}') : (project.settings || {}); } catch (e) { settings = {}; }

    var folder = resolveProjectFolder_(project);
    var docTitle = (project.name || project.id) + ' - ' + template.label;
    var doc = DocumentApp.create(docTitle);
    var body = doc.getBody();

    body.clear();

    var titlePara = body.appendParagraph(docTitle);
    titlePara.setHeading(DocumentApp.ParagraphHeading.TITLE);

    var generatedLine = 'Generated by COLONY on ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var generatedPara = body.appendParagraph(generatedLine);
    generatedPara.setItalic(true);
    generatedPara.editAsText().setForegroundColor('#666666');

    body.appendHorizontalRule();

    var metadataHeading = body.appendParagraph('Project Metadata');
    metadataHeading.setHeading(DocumentApp.ParagraphHeading.HEADING2);

    var metaRows = _buildProjectMetadataRows_(project, settings);
    if (metaRows.length) {
      var table = body.appendTable(metaRows);
      try {
        for (var r = 0; r < table.getNumRows(); r++) {
          var labelCell = table.getRow(r).getCell(0);
          labelCell.editAsText().setBold(true);
          labelCell.editAsText().setForegroundColor('#444444');
        }
      } catch (e) {}
    }

    body.appendParagraph('');

    template.sections.forEach(function(section) {
      var heading = body.appendParagraph(section.heading);
      heading.setHeading(DocumentApp.ParagraphHeading.HEADING2);
      var hint = body.appendParagraph(section.body);
      hint.setItalic(true);
      hint.editAsText().setForegroundColor('#888888');
      body.appendParagraph('');
    });

    doc.saveAndClose();

    var docId = doc.getId();
    var driveFile = DriveApp.getFileById(docId);
    var existingParents = driveFile.getParents();
    var alreadyInFolder = false;
    while (existingParents.hasNext()) {
      var parent = existingParents.next();
      if (parent.getId() === folder.getId()) { alreadyInFolder = true; break; }
    }
    if (!alreadyInFolder) {
      try { folder.addFile(driveFile); } catch (e) { console.error('generateDocFromBuiltinTemplate: addFile failed:', e); }
      try {
        var rootFolder = DriveApp.getRootFolder();
        if (rootFolder.getId() !== folder.getId()) rootFolder.removeFile(driveFile);
      } catch (e) {}
    }

    return {
      success: true,
      file: {
        id: docId,
        name: driveFile.getName(),
        url: driveFile.getUrl(),
        mimeType: driveFile.getMimeType()
      }
    };
  } catch (error) {
    console.error('generateDocFromBuiltinTemplate failed:', error);
    return { success: false, error: error.message };
  }
}

function _buildProjectMetadataRows_(project, settings) {
  var rows = [['Field', 'Value']];
  function add(label, value) {
    if (value === null || value === undefined) return;
    var s = String(value).trim();
    if (!s) return;
    rows.push([label, s]);
  }
  add('Project Name', project.name || project.id);
  add('Project ID', project.id);
  add('Status', project.status);
  add('Workstream', settings.workstream || project.workstream);
  add('Project Type', settings.projectType || project.projectType);
  add('Primary Owner', project.ownerId);
  add('Description', project.description);
  add('Tech Stack', _formatListField_(settings.techStack));
  add('Deployment Location', _formatListField_(settings.deploymentLocation));
  add('Repository', settings.githubLinks);
  add('Drive Folder', settings.googleDriveFolder);
  add('Start Date', project.startDate);
  if (settings.dataSourceExplain) add('Data Source', settings.dataSourceExplain);
  if (settings.dataCadence) add('Data Refresh', settings.dataCadence);
  if (settings.contractCurrent) add('Current Contract', settings.contractCurrent);
  return rows;
}

function _formatListField_(value) {
  if (!value) return '';
  if (Array.isArray(value)) return value.filter(Boolean).join(', ');
  if (typeof value === 'string') {
    try {
      var parsed = JSON.parse(value);
      if (Array.isArray(parsed)) return parsed.filter(Boolean).join(', ');
    } catch (e) {}
    return value;
  }
  return String(value);
}
