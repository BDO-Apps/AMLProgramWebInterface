# Power Automate workflows (end-to-end)

This guide describes the Power Automate flows required to run the full onboarding workflow from **Preparer submit** through **final Engagement Partner sign-off**, using **email routing based on the reviewer email fields captured at Step 7**.

The app controls workflow state via these `ClientOnboarding` columns:

- `WorkflowStatus` (Choice): `Draft` → `Pending Compliance` → `Pending Engagement Partner` → (`Pending Risk Partner` for High risk only) → `Closed` (with explicit return states)
- `CurrentHandler` (email/UPN string): the active reviewer
- `ComplianceReviewerEmail`, `RiskPartnerReviewerEmail`, `EngagementPartnerReviewerEmail` (email/UPN strings): assigned by the preparer in Step 7
- `OverallRiskRating` (Choice): `Low`/`Medium`/`High`
- `IsPEP` (Yes/No): PEP trigger

> Important: the app already updates `WorkflowStatus` and `CurrentHandler` when reviewers click Approve/Return/Reject. Power Automate is used to **notify**, **request approvals**, and optionally **lock permissions** by stage.

---

## Flow 0 (recommended): Common setup

Create these SharePoint views (optional but useful):

- **Pending Compliance**: `WorkflowStatus` = `Pending Compliance`
- **Pending Risk Partner**: `WorkflowStatus` = `Pending Risk Partner`
- **Pending Engagement Partner**: `WorkflowStatus` = `Pending Engagement Partner`
- **Returned to preparer**: `WorkflowStatus` = `ReturnedToPreparer`
- **Returned to compliance**: `WorkflowStatus` = `ReturnedToCompliance`
- **Returned to EP**: `WorkflowStatus` = `ReturnedToEP`
- **Rejected**: `WorkflowStatus` = `Rejected`
- **Closed**: `WorkflowStatus` = `Closed`

Also create an M365 group or security group per role if you want fallback routing when an email is blank:

- Compliance group
- Risk Partner group
- Engagement Partner group

---

## Flow 1: Submit → notify Compliance / MLRO (email-based routing)

### Trigger
- **SharePoint**: When an item is created or modified (ClientOnboarding)

### Trigger conditions (to avoid loops)
Only run when:
- `WorkflowStatus` equals `Pending Compliance`
- AND (`CurrentHandler` changed OR `WorkflowStatus` changed)

### Steps
1. **Initialize variables**
   - `caseId` = `Title`
   - `clientName` = `ClientName`
   - `handlerEmail` = `ComplianceReviewerEmail` (or `CurrentHandler` if you prefer)
2. **Condition**: if `handlerEmail` is empty
   - Set `handlerEmail` to Compliance group email (fallback)
3. **Send email (V2)**
   - To: `handlerEmail`
   - Subject: `AML Onboarding Review Required — {caseId} — {clientName}`
   - Body: include link to the SharePoint item and key risk fields (`OverallRiskRating`, `IsPEP`).
4. (Optional) **Create an Approval**
   - Type: Approve/Reject – First to respond
   - Assigned to: `handlerEmail`
   - Details: include case link + summary
5. (Optional) **Post to Teams** (Compliance channel)

### Outcome
Compliance gets notified immediately when the preparer submits.

---

## Flow 2: Compliance Approve/Return/Reject → notify next actor

> If you use the app buttons for sign-off, this flow should only **notify** based on the new `WorkflowStatus` and `CurrentHandler`.

### Trigger
- When an item is created or modified (ClientOnboarding)

### Branches

#### A) Returned to preparer
Condition:
- `WorkflowStatus` = `ReturnedToPreparer`

Actions:
- Email the preparer (use `Created` email/UPN) that corrections are required
- Include “comments” guidance: reviewers should enter notes in the app (AuditLog) so preparer sees the required fixes

#### B) Rejected
Condition:
- `WorkflowStatus` = `Rejected`

Actions:
- Email the preparer + engagement partner (optional) with rejection notice

#### C) Approved by Compliance → route to Engagement Partner
Condition:
- `WorkflowStatus` = `Pending Engagement Partner`

Actions:
- Email `EngagementPartnerReviewerEmail` (fallback to Engagement Partner group)

---

## Flow 3: Engagement Partner Approve/Return → notify Risk Partner (High risk only) or close-out

### Trigger
- When an item is created or modified (ClientOnboarding)

### Branches

#### A) Returned to Compliance
Condition:
- `WorkflowStatus` = `ReturnedToCompliance`

Actions:
- Email preparer (Created By)

#### B) Approved (Low/Medium) → final close-out
Condition:
- `WorkflowStatus` = `Closed`

Actions:
- Email preparer + compliance (optional)

#### C) Approved (High) → route to Risk Partner
Condition:
- `WorkflowStatus` = `Pending Risk Partner`

Actions:
- Email `RiskPartnerReviewerEmail` (fallback to Risk Partner group)
- Include risk reason: `OverallRiskRating` = High

---

## Flow 4: Risk Partner Approve/Return → final close-out (High risk only)

### Trigger
- When an item is created or modified (ClientOnboarding)

### Branches

#### A) Returned to Engagement Partner
Condition:
- `WorkflowStatus` = `ReturnedToEP`

Actions:
- Email preparer (Created By)

#### B) Approved → close-out
Condition:
 - `WorkflowStatus` = `Closed`

Actions:
- Email preparer + compliance (optional)

---

## Flow 5: Final close-out notifications (Closed / Rejected)

### Trigger
- When an item is created or modified (ClientOnboarding)

### Condition
- `WorkflowStatus` = `Closed`

### Actions
1. Email preparer that onboarding is approved (include `ReviewFrequency` and `NextReviewDate`)
2. Email compliance (optional) to confirm final approval
3. (Optional) Move/label documents in the library or apply retention labels

---

## Flow 6 (optional but recommended): Permissions lock by stage

Goal: after submission, prevent preparer edits in SharePoint list UI and restrict who can edit the item.

### Trigger
- When item created or modified (ClientOnboarding)

### Rules
- `Draft` / `ReturnedToPreparer`: Preparer can edit
- `Pending Compliance` / `Pending Risk Partner` / `Pending Engagement Partner`: only the active handler (and admins) can edit
- `Closed` / `Rejected`: read-only for all except admins

### Implementation approach (common pattern)
1. **Send an HTTP request to SharePoint**
   - Break role inheritance on the item
2. **Grant access to a set of users**
   - Preparer (Created By): Read or Contribute depending on status
   - Compliance reviewer email: Contribute when pending compliance
   - Risk partner reviewer email: Contribute when pending risk
   - Engagement partner reviewer email: Contribute when pending partner
3. (Optional) If you don’t want per-user permissions, use role-based M365 groups instead.

> This flow requires SharePoint admin permissions and careful testing to avoid locking out admins.

---

## Flow 7 (optional): Reminders / escalation

### Trigger
- Scheduled (Recurrence), e.g. every weekday at 08:00

### Steps
1. Get items where `WorkflowStatus` starts with `Pending`
2. For each item, compute “age” since `LastUpdated` or `DateCreated`
3. If older than threshold (e.g. 2 business days), email `CurrentHandler` and CC compliance lead

---

## Required columns for email-based routing

Ensure the `ClientOnboarding` list contains these internal names:

- `WorkflowStatus` (Choice)
- `CurrentHandler` (single line of text OR person; email/UPN value)
- `ComplianceReviewerEmail` (single line of text)
- `RiskPartnerReviewerEmail` (single line of text)
- `EngagementPartnerReviewerEmail` (single line of text)
- `OverallRiskRating` (Choice)
- `IsPEP` (Yes/No)

---

## Canonical `WorkflowStatus` map

Use these exact status values in SharePoint and Power Automate conditions:

- Preparer submits to Compliance → `Pending Compliance`
- Compliance returns to preparer → `ReturnedToPreparer`
- Compliance approves → `Pending Engagement Partner`
- Engagement Partner returns to Compliance → `ReturnedToCompliance`
- Engagement Partner approves (Medium/Low) → `Closed`
- Engagement Partner approves (High) → `Pending Risk Partner`
- Risk Partner returns to Engagement Partner → `ReturnedToEP`
- Risk Partner approves → `Closed`
- Any party rejects → `Rejected`

