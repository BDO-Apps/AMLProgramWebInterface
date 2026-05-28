import { getDefaultWorkflowAssignees } from '../../config/workflowDefaults';
import type {
  AuditLog,
  BeneficialOwner,
  Director,
  DocumentUpload,
  OnboardingCase,
} from '../SharePointService';
import type { GraphListItem } from './graphClient';
import { aliasMaps, columnSchema } from './columnSchema';
import { sharePointConfig } from '../../config/sharepointConfig';

const LOOKUP_ID_FIELD = `${sharePointConfig.caseLookupField}LookupId`;

function asString(value: unknown, fallback = ''): string {
  if (value == null) return fallback;
  return String(value);
}

function asBool(value: unknown): boolean {
  return value === true || value === 'true' || value === 'Yes';
}

function asNumber(value: unknown): number {
  const n = Number(value);
  return Number.isFinite(n) ? n : 0;
}

function formatFileSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}

function readChoice(
  fields: Record<string, unknown>,
  aliases: readonly string[],
  fallback: string
): string {
  const value = columnSchema.getFieldValue(fields, aliases);
  return value != null && value !== '' ? asString(value) : fallback;
}

export function buildCaseFolderName(caseId: string, clientName: string): string {
  const safeName = clientName.replace(/[\\/:*?"<>|]/g, '').trim();
  return `[${caseId}] ${safeName}`.slice(0, 200);
}

/**
 * Flattens nested OnboardingCase → logical SharePoint column keys.
 * Does not send to Graph until passed through columnSchema.mapFields().
 */
export function buildClientOnboardingLogicalFields(
  caseData: OnboardingCase
): Record<string, unknown> {
  const fields: Record<string, unknown> = {
    Title: caseData.id,
    ClientName: caseData.clientName,
    ClientType: caseData.clientType,
    RegNumber: caseData.regNumber,
    RegisteredAddress: caseData.registeredAddress,
    NatureOfBusiness: caseData.natureOfBusiness,
    PurposeOfEngagement: caseData.purposeOfEngagement,
    ContactAddress: caseData.contactInfo.address,
    ContactEmail: caseData.contactInfo.email,
    ContactPhone: caseData.contactInfo.phone,
    Office: caseData.office,
    HasBeneficialOwners: caseData.hasBeneficialOwners,
    RiskRatingClient: caseData.riskRatings.client,
    RiskRatingGeography: caseData.riskRatings.geography,
    RiskRatingProductService: caseData.riskRatings.productService,
    RiskRatingDeliveryChannel: caseData.riskRatings.deliveryChannel,
    RiskRatingPaymentMode: caseData.riskRatings.paymentMode,
    OverallRiskRating: caseData.overallRiskRating,
    RiskRationale: caseData.riskRationale,
    ClientRiskPEPOrAssociate: caseData.riskIndicators.client.pepOrAssociate,
    ClientRiskNPO: caseData.riskIndicators.client.npo,
    ClientRiskComplexOwnership: caseData.riskIndicators.client.complexOwnership,
    ClientRiskCashIntensive: caseData.riskIndicators.client.cashIntensive,
    ClientRiskIntermediaries: caseData.riskIndicators.client.intermediaries,
    ClientRiskOther: caseData.riskIndicators.client.other,
    GeoRiskFATFGreyListed: caseData.riskIndicators.geography.fatfGreyListed,
    GeoRiskFATFBlackListed: caseData.riskIndicators.geography.fatfBlackListed,
    GeoRiskSanctionsExposed: caseData.riskIndicators.geography.sanctionsExposed,
    GeoRiskHighCorruptionConflict: caseData.riskIndicators.geography.highCorruptionOrConflict,
    GeoRiskOther: caseData.riskIndicators.geography.other,
    ProdRiskTrustCompanyFormation: caseData.riskIndicators.productService.trustOrCompanyFormation,
    ProdRiskManageClientFundsAssets: caseData.riskIndicators.productService.manageClientFundsOrAssets,
    ProdRiskCrossBorderTransactions: caseData.riskIndicators.productService.crossBorderTransactions,
    ProdRiskHighValueComplexTransactions: caseData.riskIndicators.productService.highValueOrComplexTransactions,
    ProdRiskOther: caseData.riskIndicators.productService.other,
    CDDIdentityVerified: caseData.cddMeasures.identityVerified,
    CDDBOVerified: caseData.cddMeasures.boVerified,
    CDDNatureUnderstood: caseData.cddMeasures.natureUnderstood,
    CDDPEPScreened: caseData.cddMeasures.pepScreened,
    CDDSanctionsScreened: caseData.cddMeasures.sanctionsScreened,
    CDDAdverseMediaScreened: caseData.cddMeasures.adverseMediaScreened,
    EDDSourceOfFundsVerified: caseData.eddApplied.sourceOfFundsVerified,
    EDDSourceOfWealthVerified: caseData.eddApplied.sourceOfWealthVerified,
    EDDEnhancedAdverseMedia: caseData.eddApplied.enhancedAdverseMedia,
    EDDAdditionalBOVerification: caseData.eddApplied.additionalBoVerification,
    EDDSeniorMgmtApproved: caseData.eddApplied.seniorMgmtApproved,
    EDDEnhancedMonitoringApplied: caseData.eddApplied.enhancedMonitoringApplied,
    EDDFindings: caseData.eddFindings,
    IsPEP: caseData.pepStatus.isPep,
    SanctionsScreened: caseData.sanctionsCheck.screened,
    SanctionsHasMatch: caseData.sanctionsCheck.hasMatch,
    SanctionsDetails: caseData.sanctionsCheck.details ?? '',
    AdverseMediaHasInfo: caseData.adverseMediaCheck.hasAdverseInfo,
    AdverseMediaDetails: caseData.adverseMediaCheck.details ?? '',
    WorkflowStatus: caseData.status,
    CurrentHandler: caseData.currentHandler,
    ComplianceReviewerEmail: caseData.workflowAssignees.complianceEmail,
    EngagementPartnerReviewerEmail: caseData.workflowAssignees.engagementPartnerEmail,
    RiskPartnerReviewerEmail: caseData.workflowAssignees.riskPartnerEmail,
    WorkflowAssigneesJson: JSON.stringify(caseData.workflowAssignees),
    LastUpdated: caseData.lastUpdated,
  };

  if (caseData.decision) fields.Decision = caseData.decision;
  if (caseData.reviewFrequency) fields.ReviewFrequency = caseData.reviewFrequency;
  if (caseData.dateCreated) fields.DateCreated = caseData.dateCreated;
  if (caseData.pepStatus.pepType) fields.PEPType = caseData.pepStatus.pepType;
  if (caseData.nextReviewDate) fields.NextReviewDate = caseData.nextReviewDate;

  const sig = caseData.signatures;
  if (sig.preparer) {
    fields.PreparerSignName = sig.preparer.name;
    fields.PreparerSignDate = sig.preparer.date;
    fields.PreparerSignText = sig.preparer.sign;
  }
  if (sig.compliance) {
    fields.ComplianceSignName = sig.compliance.name;
    fields.ComplianceSignDate = sig.compliance.date;
    fields.ComplianceSignText = sig.compliance.sign;
  }
  if (sig.engagementPartner) {
    fields.EngagementPartnerSignName = sig.engagementPartner.name;
    fields.EngagementPartnerSignDate = sig.engagementPartner.date;
    fields.EngagementPartnerSignText = sig.engagementPartner.sign;
  }
  if (sig.riskPartner) {
    fields.RiskPartnerSignName = sig.riskPartner.name;
    fields.RiskPartnerSignDate = sig.riskPartner.date;
    fields.RiskPartnerSignText = sig.riskPartner.sign;
  }
  fields.ComplianceReviewComment = caseData.reviewComments.compliance;
  fields.EPReviewComment = caseData.reviewComments.engagementPartner;
  fields.RPReviewComment = caseData.reviewComments.riskPartner;

  return fields;
}

/** Maps to actual SharePoint columns present on the list (skips unknown fields). */
export async function caseToSharePointFields(
  siteId: string,
  listId: string,
  caseData: OnboardingCase
): Promise<Record<string, unknown>> {
  const logical = buildClientOnboardingLogicalFields(caseData);
  const columns = await columnSchema.loadColumns(siteId, listId);
  const columnNames = new Set(columns.map((c) => c.name));

  const hasFlatContact = (aliasMaps.clientOnboarding.ContactEmail ?? []).some((a) =>
    columnNames.has(a)
  );

  const hasContactInfoJson = (aliasMaps.clientOnboarding.ContactInfo ?? []).some((a) =>
    columnNames.has(a)
  );

  if (!hasFlatContact && hasContactInfoJson) {
    logical.ContactInfo = JSON.stringify(caseData.contactInfo);
    delete logical.ContactAddress;
    delete logical.ContactEmail;
    delete logical.ContactPhone;
  }

  const hasAnyReviewerEmailFlat = [
    ...(aliasMaps.clientOnboarding.ComplianceReviewerEmail ?? []),
    ...(aliasMaps.clientOnboarding.EngagementPartnerReviewerEmail ?? []),
    ...(aliasMaps.clientOnboarding.RiskPartnerReviewerEmail ?? []),
  ].some((a) => columnNames.has(a));

  const hasReviewerJson = (aliasMaps.clientOnboarding.WorkflowAssigneesJson ?? []).some((a) =>
    columnNames.has(a)
  );

  // If the list doesn't have the 3 flat reviewer email columns, persist as JSON fallback
  if (!hasAnyReviewerEmailFlat && hasReviewerJson) {
    logical.WorkflowAssigneesJson = JSON.stringify(caseData.workflowAssignees);
    delete logical.ComplianceReviewerEmail;
    delete logical.EngagementPartnerReviewerEmail;
    delete logical.RiskPartnerReviewerEmail;
  }

  const mapped = await columnSchema.mapFields(
    siteId,
    listId,
    logical,
    aliasMaps.clientOnboarding
  );

  const skipped = Object.keys(logical).filter(
    (k) => !(k in mapped) && logical[k] !== undefined && logical[k] !== ''
  );
  if (skipped.length > 0) {
    console.warn(
      '[SharePoint] ClientOnboarding fields not on list (skipped):',
      skipped.join(', ')
    );
  }

  return mapped;
}

export function sharePointItemToCase(
  item: GraphListItem,
  directors: Director[],
  beneficialOwners: BeneficialOwner[],
  auditLogs: AuditLog[],
  documents: DocumentUpload[]
): OnboardingCase {
  const f = item.fields;
  const A = aliasMaps.clientOnboarding;

  let contactInfo = {
    address: readChoice(f, A.ContactAddress ?? ['ContactAddress'], ''),
    email: readChoice(f, A.ContactEmail ?? ['ContactEmail'], ''),
    phone: readChoice(f, A.ContactPhone ?? ['ContactPhone'], ''),
  };

  const contactJson = columnSchema.getFieldValue(f, A.ContactInfo ?? ['ContactInfo']);
  if (contactJson && typeof contactJson === 'string') {
    try {
      const parsed = JSON.parse(contactJson) as {
        address?: string;
        email?: string;
        phone?: string;
      };
      contactInfo = {
        // Prefer flat columns when present; only fall back to JSON for blanks.
        address: contactInfo.address || parsed.address || '',
        email: contactInfo.email || parsed.email || '',
        phone: contactInfo.phone || parsed.phone || '',
      };
    } catch {
      /* keep flat fields */
    }
  }

  const handlerRaw = columnSchema.getFieldValue(
    f,
    A.CurrentHandler ?? ['CurrentHandler']
  );
  let currentHandler = '';
  if (typeof handlerRaw === 'string') currentHandler = handlerRaw;
  else if (handlerRaw && typeof handlerRaw === 'object') {
    const h = handlerRaw as Record<string, unknown>;
    currentHandler = asString(h.Email ?? h.email ?? h.Title ?? h.title);
  }

  return {
    id: readChoice(f, A.Title ?? ['Title'], ''),
    sharePointItemId: item.id,
    clientName: readChoice(f, A.ClientName ?? ['ClientName'], ''),
    clientType: readChoice(f, A.ClientType ?? ['ClientType'], 'Individual') as OnboardingCase['clientType'],
    regNumber: readChoice(f, A.RegNumber ?? ['RegNumber'], ''),
    registeredAddress: readChoice(f, A.RegisteredAddress ?? ['RegisteredAddress'], ''),
    natureOfBusiness: readChoice(f, A.NatureOfBusiness ?? ['NatureOfBusiness'], ''),
    purposeOfEngagement: readChoice(f, A.PurposeOfEngagement ?? ['PurposeOfEngagement'], ''),
    contactInfo,
    directors,
    beneficialOwners,
    hasBeneficialOwners: asBool(columnSchema.getFieldValue(f, A.HasBeneficialOwners ?? ['HasBeneficialOwners'])),
    riskRatings: {
      client: readChoice(f, A.RiskRatingClient ?? ['RiskRatingClient'], 'Low') as 'Low' | 'Medium' | 'High',
      geography: readChoice(f, A.RiskRatingGeography ?? ['RiskRatingGeography'], 'Low') as 'Low' | 'Medium' | 'High',
      productService: readChoice(f, A.RiskRatingProductService ?? ['RiskRatingProductService'], 'Low') as 'Low' | 'Medium' | 'High',
      deliveryChannel: readChoice(f, A.RiskRatingDeliveryChannel ?? ['RiskRatingDeliveryChannel'], 'Low') as 'Low' | 'Medium' | 'High',
      paymentMode: readChoice(f, A.RiskRatingPaymentMode ?? ['RiskRatingPaymentMode'], 'Low') as 'Low' | 'Medium' | 'High',
    },
    riskIndicators: {
      client: {
        pepOrAssociate: asBool(columnSchema.getFieldValue(f, A.ClientRiskPEPOrAssociate ?? ['ClientRiskPEPOrAssociate'])),
        npo: asBool(columnSchema.getFieldValue(f, A.ClientRiskNPO ?? ['ClientRiskNPO'])),
        complexOwnership: asBool(columnSchema.getFieldValue(f, A.ClientRiskComplexOwnership ?? ['ClientRiskComplexOwnership'])),
        cashIntensive: asBool(columnSchema.getFieldValue(f, A.ClientRiskCashIntensive ?? ['ClientRiskCashIntensive'])),
        intermediaries: asBool(columnSchema.getFieldValue(f, A.ClientRiskIntermediaries ?? ['ClientRiskIntermediaries'])),
        other: readChoice(f, A.ClientRiskOther ?? ['ClientRiskOther'], ''),
      },
      geography: {
        fatfGreyListed: asBool(columnSchema.getFieldValue(f, A.GeoRiskFATFGreyListed ?? ['GeoRiskFATFGreyListed'])),
        fatfBlackListed: asBool(columnSchema.getFieldValue(f, A.GeoRiskFATFBlackListed ?? ['GeoRiskFATFBlackListed'])),
        sanctionsExposed: asBool(columnSchema.getFieldValue(f, A.GeoRiskSanctionsExposed ?? ['GeoRiskSanctionsExposed'])),
        highCorruptionOrConflict: asBool(columnSchema.getFieldValue(f, A.GeoRiskHighCorruptionConflict ?? ['GeoRiskHighCorruptionConflict'])),
        other: readChoice(f, A.GeoRiskOther ?? ['GeoRiskOther'], ''),
      },
      productService: {
        trustOrCompanyFormation: asBool(columnSchema.getFieldValue(f, A.ProdRiskTrustCompanyFormation ?? ['ProdRiskTrustCompanyFormation'])),
        manageClientFundsOrAssets: asBool(columnSchema.getFieldValue(f, A.ProdRiskManageClientFundsAssets ?? ['ProdRiskManageClientFundsAssets'])),
        crossBorderTransactions: asBool(columnSchema.getFieldValue(f, A.ProdRiskCrossBorderTransactions ?? ['ProdRiskCrossBorderTransactions'])),
        highValueOrComplexTransactions: asBool(columnSchema.getFieldValue(f, A.ProdRiskHighValueComplexTransactions ?? ['ProdRiskHighValueComplexTransactions'])),
        other: readChoice(f, A.ProdRiskOther ?? ['ProdRiskOther'], ''),
      },
    },
    overallRiskRating: readChoice(f, A.OverallRiskRating ?? ['OverallRiskRating'], 'Low') as 'Low' | 'Medium' | 'High',
    riskRationale: readChoice(f, A.RiskRationale ?? ['RiskRationale'], ''),
    cddMeasures: {
      identityVerified: asBool(columnSchema.getFieldValue(f, A.CDDIdentityVerified ?? ['CDDIdentityVerified'])),
      boVerified: asBool(columnSchema.getFieldValue(f, A.CDDBOVerified ?? ['CDDBOVerified'])),
      natureUnderstood: asBool(columnSchema.getFieldValue(f, A.CDDNatureUnderstood ?? ['CDDNatureUnderstood'])),
      pepScreened: asBool(columnSchema.getFieldValue(f, A.CDDPEPScreened ?? ['CDDPEPScreened'])),
      sanctionsScreened: asBool(columnSchema.getFieldValue(f, A.CDDSanctionsScreened ?? ['CDDSanctionsScreened'])),
      adverseMediaScreened: asBool(columnSchema.getFieldValue(f, A.CDDAdverseMediaScreened ?? ['CDDAdverseMediaScreened'])),
    },
    eddApplied: {
      sourceOfFundsVerified: asBool(columnSchema.getFieldValue(f, A.EDDSourceOfFundsVerified ?? ['EDDSourceOfFundsVerified'])),
      sourceOfWealthVerified: asBool(columnSchema.getFieldValue(f, A.EDDSourceOfWealthVerified ?? ['EDDSourceOfWealthVerified'])),
      enhancedAdverseMedia: asBool(columnSchema.getFieldValue(f, A.EDDEnhancedAdverseMedia ?? ['EDDEnhancedAdverseMedia'])),
      additionalBoVerification: asBool(columnSchema.getFieldValue(f, A.EDDAdditionalBOVerification ?? ['EDDAdditionalBOVerification'])),
      seniorMgmtApproved: asBool(columnSchema.getFieldValue(f, A.EDDSeniorMgmtApproved ?? ['EDDSeniorMgmtApproved'])),
      enhancedMonitoringApplied: asBool(columnSchema.getFieldValue(f, A.EDDEnhancedMonitoringApplied ?? ['EDDEnhancedMonitoringApplied'])),
    },
    eddFindings: readChoice(f, A.EDDFindings ?? ['EDDFindings'], ''),
    pepStatus: {
      isPep: asBool(columnSchema.getFieldValue(f, A.IsPEP ?? ['IsPEP'])),
      pepType: columnSchema.getFieldValue(f, A.PEPType ?? ['PEPType'])
        ? (readChoice(f, A.PEPType ?? ['PEPType'], '') as 'Domestic' | 'Foreign' | 'International')
        : undefined,
    },
    sanctionsCheck: {
      screened: asBool(columnSchema.getFieldValue(f, A.SanctionsScreened ?? ['SanctionsScreened'])),
      hasMatch: asBool(columnSchema.getFieldValue(f, A.SanctionsHasMatch ?? ['SanctionsHasMatch'])),
      details: readChoice(f, A.SanctionsDetails ?? ['SanctionsDetails'], '') || undefined,
    },
    adverseMediaCheck: {
      hasAdverseInfo: asBool(columnSchema.getFieldValue(f, A.AdverseMediaHasInfo ?? ['AdverseMediaHasInfo'])),
      details: readChoice(f, A.AdverseMediaDetails ?? ['AdverseMediaDetails'], '') || undefined,
    },
    decision: readChoice(f, A.Decision ?? ['Decision'], '') as OnboardingCase['decision'],
    reviewFrequency: readChoice(f, A.ReviewFrequency ?? ['ReviewFrequency'], '') as OnboardingCase['reviewFrequency'],
    nextReviewDate: readChoice(f, A.NextReviewDate ?? ['NextReviewDate'], '').split('T')[0] ?? '',
    signatures: {
      preparer: columnSchema.getFieldValue(f, A.PreparerSignName ?? ['PreparerSignName'])
        ? {
            name: readChoice(f, A.PreparerSignName ?? ['PreparerSignName'], ''),
            date: readChoice(f, A.PreparerSignDate ?? ['PreparerSignDate'], '').split('T')[0],
            sign: readChoice(f, A.PreparerSignText ?? ['PreparerSignText'], ''),
          }
        : undefined,
      compliance: columnSchema.getFieldValue(f, A.ComplianceSignName ?? ['ComplianceSignName'])
        ? {
            name: readChoice(f, A.ComplianceSignName ?? ['ComplianceSignName'], ''),
            date: readChoice(f, A.ComplianceSignDate ?? ['ComplianceSignDate'], '').split('T')[0],
            sign: readChoice(f, A.ComplianceSignText ?? ['ComplianceSignText'], ''),
          }
        : undefined,
      engagementPartner: columnSchema.getFieldValue(f, A.EngagementPartnerSignName ?? ['EngagementPartnerSignName'])
        ? {
            name: readChoice(f, A.EngagementPartnerSignName ?? ['EngagementPartnerSignName'], ''),
            date: readChoice(f, A.EngagementPartnerSignDate ?? ['EngagementPartnerSignDate'], '').split('T')[0],
            sign: readChoice(f, A.EngagementPartnerSignText ?? ['EngagementPartnerSignText'], ''),
          }
        : undefined,
      riskPartner: columnSchema.getFieldValue(f, A.RiskPartnerSignName ?? ['RiskPartnerSignName'])
        ? {
            name: readChoice(f, A.RiskPartnerSignName ?? ['RiskPartnerSignName'], ''),
            date: readChoice(f, A.RiskPartnerSignDate ?? ['RiskPartnerSignDate'], '').split('T')[0],
            sign: readChoice(f, A.RiskPartnerSignText ?? ['RiskPartnerSignText'], ''),
          }
        : undefined,
    },
    reviewComments: {
      compliance: readChoice(f, A.ComplianceReviewComment ?? ['ComplianceReviewComment'], ''),
      engagementPartner: readChoice(f, A.EPReviewComment ?? ['EPReviewComment'], ''),
      riskPartner: readChoice(f, A.RPReviewComment ?? ['RPReviewComment'], ''),
    },
    status: readChoice(f, A.WorkflowStatus ?? ['WorkflowStatus'], 'Draft') as OnboardingCase['status'],
    currentHandler,
    workflowAssignees: (() => {
      const fromFlat = {
        complianceEmail: readChoice(
          f,
          A.ComplianceReviewerEmail ?? ['ComplianceReviewerEmail'],
          ''
        ).toLowerCase(),
        engagementPartnerEmail: readChoice(
          f,
          A.EngagementPartnerReviewerEmail ?? ['EngagementPartnerReviewerEmail'],
          ''
        ).toLowerCase(),
        riskPartnerEmail: readChoice(
          f,
          A.RiskPartnerReviewerEmail ?? ['RiskPartnerReviewerEmail'],
          ''
        ).toLowerCase(),
      };

      const raw = columnSchema.getFieldValue(
        f,
        A.WorkflowAssigneesJson ?? ['WorkflowAssigneesJson']
      );
      if (raw && typeof raw === 'string') {
        try {
          const parsed = JSON.parse(raw) as Partial<OnboardingCase['workflowAssignees']>;
          return {
            complianceEmail: (fromFlat.complianceEmail || parsed.complianceEmail || '').toLowerCase(),
            engagementPartnerEmail: (fromFlat.engagementPartnerEmail || parsed.engagementPartnerEmail || '').toLowerCase(),
            riskPartnerEmail: (fromFlat.riskPartnerEmail || parsed.riskPartnerEmail || '').toLowerCase(),
          };
        } catch {
          /* ignore */
        }
      }

      // If nothing is stored (flat columns missing), keep safe defaults for UX.
      if (!fromFlat.complianceEmail && !fromFlat.engagementPartnerEmail && !fromFlat.riskPartnerEmail) {
        return getDefaultWorkflowAssignees();
      }

      return {
        complianceEmail: fromFlat.complianceEmail || getDefaultWorkflowAssignees().complianceEmail,
        engagementPartnerEmail:
          fromFlat.engagementPartnerEmail || getDefaultWorkflowAssignees().engagementPartnerEmail,
        riskPartnerEmail: fromFlat.riskPartnerEmail || getDefaultWorkflowAssignees().riskPartnerEmail,
      };
    })(),
    documents,
    auditLogs,
    office: readChoice(f, A.Office ?? ['Office'], 'Zimbabwe') as 'Zimbabwe' | 'Malawi',
    dateCreated: readChoice(f, A.DateCreated ?? ['DateCreated'], '').split('T')[0] ?? '',
    lastUpdated: readChoice(f, A.LastUpdated ?? ['LastUpdated'], '').split('T')[0] ?? '',
  };
}

export async function directorToFields(
  siteId: string,
  listId: string,
  director: Director,
  caseLookupId: number
): Promise<Record<string, unknown>> {
  const logical = {
    FullName: director.fullName,
    Position: director.position,
    Nationality: director.nationality,
    IDNumber: director.idNumber,
    CountryOfResidence: director.countryOfResidence,
    CaseIDLookupId: caseLookupId,
  };
  return columnSchema.mapFields(siteId, listId, logical, aliasMaps.directors);
}

export function directorFromItem(item: GraphListItem): Director {
  const f = item.fields;
  const A = aliasMaps.directors;
  return {
    id: item.id,
    fullName: readChoice(f, A.FullName ?? ['FullName'], ''),
    position: readChoice(f, A.Position ?? ['Position'], ''),
    nationality: readChoice(f, A.Nationality ?? ['Nationality'], ''),
    idNumber: readChoice(f, A.IDNumber ?? ['IDNumber'], ''),
    countryOfResidence: readChoice(f, A.CountryOfResidence ?? ['CountryOfResidence'], ''),
  };
}

export async function beneficialOwnerToFields(
  siteId: string,
  listId: string,
  bo: BeneficialOwner,
  caseLookupId: number
): Promise<Record<string, unknown>> {
  const logical = {
    FullName: bo.fullName,
    OwnershipPercentage: bo.ownershipPercentage,
    BasisOfControl: bo.basisOfControl,
    Country: bo.country,
    VerificationSource: bo.verificationSource,
    CaseIDLookupId: caseLookupId,
  };
  return columnSchema.mapFields(siteId, listId, logical, aliasMaps.beneficialOwners);
}

export function beneficialOwnerFromItem(item: GraphListItem): BeneficialOwner {
  const f = item.fields;
  const A = aliasMaps.beneficialOwners;
  return {
    id: item.id,
    fullName: readChoice(f, A.FullName ?? ['FullName'], ''),
    ownershipPercentage: asNumber(columnSchema.getFieldValue(f, A.OwnershipPercentage ?? ['OwnershipPercentage'])),
    basisOfControl: readChoice(f, A.BasisOfControl ?? ['BasisOfControl'], ''),
    country: readChoice(f, A.Country ?? ['Country'], ''),
    verificationSource: readChoice(f, A.VerificationSource ?? ['VerificationSource'], ''),
  };
}

export async function auditLogToFields(
  siteId: string,
  listId: string,
  log: AuditLog,
  caseLookupId: number
): Promise<Record<string, unknown>> {
  const logical = {
    EventTimestamp: log.timestamp,
    Actor: log.actor,
    Role: log.role,
    Action: log.action,
    Comments: log.comments,
    CaseIDLookupId: caseLookupId,
  };
  return columnSchema.mapFields(siteId, listId, logical, aliasMaps.auditLogs);
}

export function auditLogFromItem(item: GraphListItem): AuditLog {
  const f = item.fields;
  const A = aliasMaps.auditLogs;
  return {
    id: item.id,
    timestamp: readChoice(f, A.EventTimestamp ?? ['EventTimestamp', 'Timestamp'], ''),
    actor: readChoice(f, A.Actor ?? ['Actor'], ''),
    role: readChoice(f, A.Role ?? ['Role'], ''),
    action: readChoice(f, A.Action ?? ['Action'], ''),
    comments: readChoice(f, A.Comments ?? ['Comments'], ''),
  };
}

export async function documentMetadataToFields(
  siteId: string,
  listId: string,
  metadata: Record<string, unknown>
): Promise<Record<string, unknown>> {
  return columnSchema.mapFields(siteId, listId, metadata, aliasMaps.documents);
}

export function driveFileToDocument(
  file: {
    id: string;
    name: string;
    size: number;
    createdDateTime: string;
    webUrl?: string;
    fields: Record<string, unknown>;
  }
): DocumentUpload {
  const A = aliasMaps.documents;
  const category = readChoice(file.fields, A.DocumentCategory ?? ['DocumentCategory'], 'ID_Passport');
  return {
    id: file.id,
    name: file.name,
    category: category as DocumentUpload['category'],
    size: formatFileSize(file.size),
    uploadDate: file.createdDateTime.split('T')[0],
    uploadedBy: readChoice(file.fields, A.UploadedBy ?? ['UploadedBy'], ''),
    webUrl: file.webUrl,
  };
}

export function getCaseLookupIdFromChildFields(
  fields: Record<string, unknown>
): number | null {
  for (const [key, value] of Object.entries(fields)) {
    if (!key.toLowerCase().includes('lookupid')) continue;
    const n = Number(value);
    if (!Number.isNaN(n) && n > 0) return n;
  }

  const caseField = fields.CaseID ?? fields.caseID;
  if (caseField && typeof caseField === 'object' && !Array.isArray(caseField)) {
    const obj = caseField as Record<string, unknown>;
    const lookupId = obj.LookupId ?? obj.lookupId ?? obj.Id ?? obj.id;
    const n = Number(lookupId);
    if (!Number.isNaN(n) && n > 0) return n;
  }

  const aliases = [
    LOOKUP_ID_FIELD,
    'CaseIDLookupId',
    'CaseID_x003a_IDLookupId',
    'CaseIDId',
  ];
  const id = columnSchema.getFieldValue(fields, aliases);
  if (typeof id === 'number' && id > 0) return id;
  if (typeof id === 'string' && id) {
    const n = Number(id);
    if (!Number.isNaN(n) && n > 0) return n;
  }
  return null;
}
