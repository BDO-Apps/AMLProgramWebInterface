import { sharePointConfig } from '../../config/sharepointConfig';
import type { WorkflowAssignees } from '../../config/workflowDefaults';
import type {
  AuditLog,
  BeneficialOwner,
  Director,
  DocumentUpload,
  OnboardingCase,
} from '../SharePointService';
import { graphClient } from './graphClient';
import {
  auditLogFromItem,
  auditLogToFields,
  beneficialOwnerFromItem,
  beneficialOwnerToFields,
  buildCaseFolderName,
  caseToSharePointFields,
  directorFromItem,
  directorToFields,
  documentMetadataToFields,
  driveFileToDocument,
  getCaseLookupIdFromChildFields,
  sharePointItemToCase,
} from './fieldMapper';
import { listResolver, type ResolvedListIds } from './listResolver';

interface SpContext {
  siteId: string;
  driveId: string;
  lists: ResolvedListIds;
}

export class SharePointRepository {
  private async resolveContext(): Promise<SpContext> {
    const siteId = await graphClient.getSiteId(
      sharePointConfig.hostname,
      sharePointConfig.sitePath
    );
    const lists = await listResolver.resolveAll(siteId);
    const driveId = await graphClient.getDriveId(siteId, lists.documentLibrary);
    return { siteId, driveId, lists };
  }

  private async getCaseLookupId(
    ctx: SpContext,
    caseTitle: string
  ): Promise<number> {
    const item = await graphClient.getItemByTitle(
      ctx.siteId,
      ctx.lists.clientOnboarding,
      caseTitle
    );
    if (!item) throw new Error(`Case not found in SharePoint: ${caseTitle}`);
    return Number(item.id);
  }

  private groupChildrenByCaseLookup<T>(
    items: { id: string; fields: Record<string, unknown> }[],
    mapFn: (item: { id: string; fields: Record<string, unknown> }) => T
  ): Map<number, T[]> {
    const map = new Map<number, T[]>();
    for (const item of items) {
      const lookupId = getCaseLookupIdFromChildFields(item.fields);
      if (lookupId == null) continue;
      const list = map.get(lookupId) ?? [];
      list.push(mapFn(item));
      map.set(lookupId, list);
    }
    return map;
  }

  private async loadRelatedData(ctx: SpContext) {
    const [directorItems, boItems, logItems] = await Promise.all([
      graphClient.listAllItems(ctx.siteId, ctx.lists.directors),
      graphClient.listAllItems(ctx.siteId, ctx.lists.beneficialOwners),
      graphClient.listAllItems(ctx.siteId, ctx.lists.auditLogs),
    ]);

    return {
      directorsByCase: this.groupChildrenByCaseLookup(directorItems, directorFromItem),
      bosByCase: this.groupChildrenByCaseLookup(boItems, beneficialOwnerFromItem),
      logsByCase: this.groupChildrenByCaseLookup(logItems, auditLogFromItem),
    };
  }

  private async loadDocumentsForCase(
    ctx: SpContext,
    caseData: OnboardingCase
  ): Promise<DocumentUpload[]> {
    const folder = buildCaseFolderName(caseData.id, caseData.clientName);
    const files = await graphClient.listFilesInFolder(
      ctx.siteId,
      ctx.driveId,
      folder
    );
    return files.map(driveFileToDocument);
  }

  async getCases(role: string, email: string): Promise<OnboardingCase[]> {
    const ctx = await this.resolveContext();
    const { directorsByCase, bosByCase, logsByCase } =
      await this.loadRelatedData(ctx);

    const caseItems = await graphClient.listAllItems(
      ctx.siteId,
      ctx.lists.clientOnboarding
    );

    const cases: OnboardingCase[] = [];

    for (const item of caseItems) {
      const lookupId = Number(item.id);
      const onboardingCase = sharePointItemToCase(
        item,
        directorsByCase.get(lookupId) ?? [],
        bosByCase.get(lookupId) ?? [],
        logsByCase.get(lookupId) ?? [],
        []
      );
      onboardingCase.documents = await this.loadDocumentsForCase(ctx, onboardingCase);
      cases.push(onboardingCase);
    }

    const normalizedEmail = email.toLowerCase();

    if (role === 'Preparer') {
      return cases.filter(
        (c) =>
          c.status === 'Draft' ||
          c.status === 'ReturnedToPreparer' ||
          c.currentHandler.toLowerCase() === normalizedEmail ||
          c.auditLogs.some((log) => log.actor.toLowerCase() === normalizedEmail)
      );
    }

    if (role === 'RiskPartner') {
      return cases.filter((c) => c.overallRiskRating === 'High');
    }

    return cases;
  }

  async getCaseById(caseId: string): Promise<OnboardingCase | null> {
    const ctx = await this.resolveContext();
    const item = await graphClient.getItemByTitle(
      ctx.siteId,
      ctx.lists.clientOnboarding,
      caseId
    );
    if (!item) return null;

    const lookupId = Number(item.id);
    const { directorsByCase, bosByCase, logsByCase } =
      await this.loadRelatedData(ctx);

    const onboardingCase = sharePointItemToCase(
      item,
      directorsByCase.get(lookupId) ?? [],
      bosByCase.get(lookupId) ?? [],
      logsByCase.get(lookupId) ?? [],
      []
    );
    onboardingCase.documents = await this.loadDocumentsForCase(ctx, onboardingCase);
    return onboardingCase;
  }

  private async syncDirectors(
    ctx: SpContext,
    caseLookupId: number,
    directors: Director[],
    existingItems: { id: string; fields: Record<string, unknown> }[]
  ) {
    const toDelete = existingItems.filter(
      (item) => getCaseLookupIdFromChildFields(item.fields) === caseLookupId
    );
    await Promise.all(
      toDelete.map((item) =>
        graphClient.deleteListItem(ctx.siteId, ctx.lists.directors, item.id)
      )
    );
    for (const director of directors) {
      const fields = await directorToFields(
        ctx.siteId,
        ctx.lists.directors,
        director,
        caseLookupId
      );
      console.info('[SharePoint] Create Director with Case lookup:', fields);
      await graphClient.createListItem(ctx.siteId, ctx.lists.directors, fields);
    }
  }

  private async syncBeneficialOwners(
    ctx: SpContext,
    caseLookupId: number,
    owners: BeneficialOwner[],
    existingItems: { id: string; fields: Record<string, unknown> }[]
  ) {
    const toDelete = existingItems.filter(
      (item) => getCaseLookupIdFromChildFields(item.fields) === caseLookupId
    );
    await Promise.all(
      toDelete.map((item) =>
        graphClient.deleteListItem(ctx.siteId, ctx.lists.beneficialOwners, item.id)
      )
    );
    for (const bo of owners) {
      await graphClient.createListItem(
        ctx.siteId,
        ctx.lists.beneficialOwners,
        await beneficialOwnerToFields(ctx.siteId, ctx.lists.beneficialOwners, bo, caseLookupId)
      );
    }
  }

  private async appendAuditLog(ctx: SpContext, caseLookupId: number, log: AuditLog) {
    await graphClient.createListItem(
      ctx.siteId,
      ctx.lists.auditLogs,
      await auditLogToFields(ctx.siteId, ctx.lists.auditLogs, log, caseLookupId)
    );
  }

  async generateCaseId(ctx: SpContext): Promise<string> {
    const year = new Date().getFullYear();
    const prefix = `BDO-AML-${year}-`;
    const items = await graphClient.listAllItems(
      ctx.siteId,
      ctx.lists.clientOnboarding
    );
    const numbers = items
      .map((item) => asString(item.fields.Title))
      .filter((title) => title.startsWith(prefix))
      .map((title) => parseInt(title.replace(prefix, ''), 10))
      .filter((n) => !Number.isNaN(n));

    const next = numbers.length > 0 ? Math.max(...numbers) + 1 : 1;
    return `${prefix}${String(next).padStart(4, '0')}`;
  }

  async saveCase(
    caseData: OnboardingCase,
    actor: string,
    role: string
  ): Promise<OnboardingCase> {
    const ctx = await this.resolveContext();
    const cleanCase = { ...caseData };
    cleanCase.lastUpdated = new Date().toISOString().split('T')[0];

    const existing = cleanCase.id
      ? await graphClient.getItemByTitle(
          ctx.siteId,
          ctx.lists.clientOnboarding,
          cleanCase.id
        )
      : null;

    const fields = await caseToSharePointFields(
      ctx.siteId,
      ctx.lists.clientOnboarding,
      cleanCase
    );
    let caseLookupId: number;
    let logAction: AuditLog['action'] = 'CREATED';
    let logComment = 'New onboarding case record established';

    if (existing) {
      caseLookupId = Number(existing.id);
      await graphClient.updateListItem(
        ctx.siteId,
        ctx.lists.clientOnboarding,
        existing.id,
        fields
      );
      logAction = 'UPDATED';
      logComment = 'Client details updated by Preparer';
    } else {
      if (!cleanCase.id) {
        cleanCase.id = await this.generateCaseId(ctx);
        fields.Title = cleanCase.id;
      }
      cleanCase.dateCreated = cleanCase.dateCreated || cleanCase.lastUpdated;
      fields.DateCreated = cleanCase.dateCreated;

      const created = await graphClient.createListItem(
        ctx.siteId,
        ctx.lists.clientOnboarding,
        fields
      );
      caseLookupId = Number(created.id);
    }

    const [directorItems, boItems] = await Promise.all([
      graphClient.listAllItems(ctx.siteId, ctx.lists.directors),
      graphClient.listAllItems(ctx.siteId, ctx.lists.beneficialOwners),
    ]);

    await this.syncDirectors(ctx, caseLookupId, cleanCase.directors, directorItems);
    await this.syncBeneficialOwners(ctx, caseLookupId, cleanCase.beneficialOwners, boItems);

    await this.appendAuditLog(ctx, caseLookupId, {
      id: `log-${Date.now()}`,
      timestamp: new Date().toISOString(),
      actor,
      role,
      action: logAction,
      comments: logComment,
    });

    const saved = await this.getCaseById(cleanCase.id);
    if (!saved) throw new Error('Failed to reload case after save');
    return saved;
  }

  async submitCase(
    caseId: string,
    actor: string,
    role: string,
    preparerSign: string,
    assignees?: WorkflowAssignees
  ): Promise<OnboardingCase> {
    const existing = await this.getCaseById(caseId);
    if (!existing) throw new Error('Case not found');

    if (assignees) existing.workflowAssignees = assignees;
    existing.status = 'Pending Compliance';
    existing.currentHandler = existing.workflowAssignees.complianceEmail.toLowerCase();
    existing.signatures.preparer = {
      name: actor,
      date: new Date().toISOString().split('T')[0],
      sign: preparerSign,
    };
    existing.lastUpdated = new Date().toISOString().split('T')[0];

    const ctx = await this.resolveContext();
    const caseLookupId = await this.getCaseLookupId(ctx, caseId);

    await graphClient.updateListItem(
      ctx.siteId,
      ctx.lists.clientOnboarding,
      String(caseLookupId),
      await caseToSharePointFields(ctx.siteId, ctx.lists.clientOnboarding, existing)
    );

    await this.appendAuditLog(ctx, caseLookupId, {
      id: `log-${Date.now()}`,
      timestamp: new Date().toISOString(),
      actor,
      role,
      action: 'SUBMITTED',
      comments:
        'Form finalized and locked for review. Mandatory documentation successfully attached.',
    });

    const updated = await this.getCaseById(caseId);
    if (!updated) throw new Error('Failed to reload case after submit');
    return updated;
  }

  async processWorkflowAction(
    caseId: string,
    action: 'Approve' | 'Return' | 'Reject',
    comments: string,
    actor: string,
    role: string,
    signatureSign: string
  ): Promise<OnboardingCase> {
    const targetCase = await this.getCaseById(caseId);
    if (!targetCase) throw new Error('Case not found');

    const timestamp = new Date().toISOString().split('T')[0];
    let nextStatus = targetCase.status;
    let nextHandler = targetCase.currentHandler;

    if (action === 'Reject') {
      nextStatus = 'Rejected';
      nextHandler = '';
      targetCase.decision = 'Client declined';
    } else if (action === 'Return') {
      if (role === 'Compliance') {
        nextStatus = 'ReturnedToPreparer';
        nextHandler = targetCase.auditLogs.find((l) => l.action === 'CREATED')?.actor ?? '';
      } else if (role === 'EngagementPartner') {
        nextStatus = 'ReturnedToCompliance';
        nextHandler = targetCase.workflowAssignees.complianceEmail.toLowerCase();
      } else if (role === 'RiskPartner') {
        nextStatus = 'ReturnedToEP';
        nextHandler = targetCase.workflowAssignees.engagementPartnerEmail.toLowerCase();
      }
    } else if (role === 'Compliance') {
      targetCase.reviewComments.compliance = comments.trim();
      targetCase.signatures.compliance = {
        name: actor,
        date: timestamp,
        sign: signatureSign,
      };
      nextStatus = 'Pending Engagement Partner';
      nextHandler = targetCase.workflowAssignees.engagementPartnerEmail.toLowerCase();
    } else if (role === 'RiskPartner') {
      targetCase.reviewComments.riskPartner = comments.trim();
      targetCase.signatures.riskPartner = {
        name: actor,
        date: timestamp,
        sign: signatureSign,
      };
      nextStatus = 'Closed';
      nextHandler = '';

      // Final close-out (only occurs after Risk Partner on High risk)
      targetCase.decision = 'Client accepted subject to conditions/EDD';

      const baseDate = new Date();
      baseDate.setMonth(baseDate.getMonth() + 6);
      targetCase.reviewFrequency = 'Enhanced and continuous';
      targetCase.nextReviewDate = baseDate.toISOString().split('T')[0];
    } else if (role === 'EngagementPartner') {
      targetCase.reviewComments.engagementPartner = comments.trim();
      targetCase.signatures.engagementPartner = {
        name: actor,
        date: timestamp,
        sign: signatureSign,
      };
      if (targetCase.overallRiskRating === 'High') {
        // High risk invokes Risk Partner AFTER Engagement Partner
        nextStatus = 'Pending Risk Partner';
        nextHandler = targetCase.workflowAssignees.riskPartnerEmail.toLowerCase();
      } else {
        // Medium/Low finalizes here
        nextStatus = 'Closed';
        nextHandler = '';
        targetCase.decision = 'Client accepted';

        const baseDate = new Date();
        if (targetCase.overallRiskRating === 'Medium') {
          baseDate.setFullYear(baseDate.getFullYear() + 1);
          targetCase.reviewFrequency = 'Annual';
        } else {
          baseDate.setFullYear(baseDate.getFullYear() + 2);
          targetCase.reviewFrequency = 'Periodic';
        }
        targetCase.nextReviewDate = baseDate.toISOString().split('T')[0];
      }
    }

    targetCase.status = nextStatus;
    targetCase.currentHandler = nextHandler;
    targetCase.lastUpdated = timestamp;

    const ctx = await this.resolveContext();
    const caseLookupId = await this.getCaseLookupId(ctx, caseId);

    await graphClient.updateListItem(
      ctx.siteId,
      ctx.lists.clientOnboarding,
      String(caseLookupId),
      await caseToSharePointFields(ctx.siteId, ctx.lists.clientOnboarding, targetCase)
    );

    await this.appendAuditLog(ctx, caseLookupId, {
      id: `log-${Date.now()}`,
      timestamp: new Date().toISOString(),
      actor,
      role,
      action: action.toUpperCase(),
      comments: comments || `${role} executed ${action.toLowerCase()} sign-off.`,
    });

    const updated = await this.getCaseById(caseId);
    if (!updated) throw new Error('Failed to reload case after workflow action');
    return updated;
  }

  async uploadDocument(
    caseId: string,
    file: File,
    category: DocumentUpload['category'],
    actor: string
  ): Promise<DocumentUpload> {
    const targetCase = await this.getCaseById(caseId);
    if (!targetCase) throw new Error('Case not found');

    const ctx = await this.resolveContext();
    const folderPath = buildCaseFolderName(caseId, targetCase.clientName);

    const docMetadata = await documentMetadataToFields(
      ctx.siteId,
      ctx.lists.documentLibrary,
      {
        CaseID: caseId,
        DocumentCategory: category,
        UploadedBy: actor,
        UploadDate: new Date().toISOString(),
      }
    );

    const driveItem = await graphClient.uploadFileToFolder(
      ctx.siteId,
      ctx.driveId,
      folderPath,
      file.name,
      file,
      docMetadata
    );

    const caseLookupId = await this.getCaseLookupId(ctx, caseId);
    await this.appendAuditLog(ctx, caseLookupId, {
      id: `log-${Date.now()}`,
      timestamp: new Date().toISOString(),
      actor,
      role: 'Preparer',
      action: 'DOCUMENT_UPLOADED',
      comments: `Uploaded document file: ${file.name} classified under ${category}.`,
    });

    await graphClient.updateListItem(
      ctx.siteId,
      ctx.lists.clientOnboarding,
      String(caseLookupId),
      { LastUpdated: new Date().toISOString().split('T')[0] }
    );

    return {
      id: driveItem.id,
      name: file.name,
      category,
      size: `${(file.size / 1024 / 1024).toFixed(1)} MB`,
      uploadDate: new Date().toISOString().split('T')[0],
      uploadedBy: actor,
      webUrl: driveItem.webUrl,
    };
  }

  async diagnoseColumns(): Promise<Array<{ name: string; displayName: string; type: string }>> {
    const ctx = await this.resolveContext();
    return graphClient.getListColumns(ctx.siteId, ctx.lists.clientOnboarding);
  }
}

function asString(value: unknown, fallback = ''): string {
  if (value == null) return fallback;
  return String(value);
}

export const sharePointRepository = new SharePointRepository();
