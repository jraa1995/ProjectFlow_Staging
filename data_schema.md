## Ad Hoc Request Schema

| Field | Type | Description |
|-------|------|-------------|
| requestId | String | Unique ID (e.g., REQ-2026-000001) |
| requestName | String | Title/subject of request |
| requestor | String | Email of person making request |
| dateOfRequest | DateTime | When request was created |
| descriptionOfRequest | String | Detailed description |
| businessJustification | String | Why it's needed |
| requestTypeCategory | Enum | Data Pull, Report Generation, System Configuration, Minor Bug Fix, Content Update, Training Material, Documentation, Analysis, Process Improvement, Other |
| whatContractDoesItApplyTo | Enum | SQuAT, Forward |
| taskName | String | Associated task/program |
| status | Enum | new, assigned, in-progress, pending-clarification, on-hold, in-review, completed, resolved, closed, cancelled, rejected |
| priority | Enum | Urgent, High, Medium, Low, Routine |
| assignedTeam | String | Team responsible |
| assignedTo | String | Assigned person's email |
| assignedToName | String | Assigned person's display name |
| dateAssigned | DateTime | When request was assigned |
| contractorPoc | String | Contractor point of contact |
| contractorPocEmail | String | Contractor POC email |
| targetCompletionDate | Date | Expected completion |
| dueDeadline | Date | Hard deadline |
| timeToCompleteEstimatedHours | Number | Estimated hours |
| timeToCompleteActualHours | Number | Actual hours spent |
| dateDeliveredCompleted | DateTime | When work was completed |
| startDateTime | DateTime | When work started |
| completionDateTime | DateTime | When work completed |
| deliverables | Array | [{name, id, url, mimeType, size}] |
| attachments | Array | [{name, id, url, mimeType, size}] |
| internalNotes | String | Internal notes (team only) |
| resolutionNotes | String | Final resolution notes |
| resolutionDate | DateTime | When resolved |
| satisfactionRating | String | Requestor feedback |
| slaBreached | Boolean | Whether SLA was violated |
| escalated | Boolean | Whether request is escalated |
| notificationEmails | String | Comma-separated notification emails |
| topic | String | Topic/category |
| difficultyLevel | Enum | Easy, Medium, Hard |
| createdBy | String | Creator's email |
| createdDate | DateTime | Creation timestamp |
| lastUpdatedDate | DateTime | Last update timestamp |
| lastUpdatedBy | String | Last updater's email |
| rowVersion | Number | Version for optimistic locking |

---

## Activity Log (Comments/Updates)

| Field | Type | Description |
|-------|------|-------------|
| logId | String | Unique ID (e.g., LOG-1736692200000) |
| requestId | String | Parent request ID |
| timestamp | DateTime | When activity occurred |
| userId | String | User's email |
| userName | String | User's display name |
| actionType | Enum | create, update, delete, status-change, comment |
| actionDetails | String | Comment text or action description (attachments embedded as markdown: [filename](url)) |
| oldValue | JSON String | Previous values (for updates) |
| newValue | JSON String | New values (for updates) |

---

## Attachment Object (used in attachments, deliverables, and embedded in comments)

| Field | Type | Description |
|-------|------|-------------|
| name | String | Filename |
| id | String | Google Drive file ID |
| url | String | Google Drive URL |
| mimeType | String | File MIME type |
| size | Number | File size in bytes |




